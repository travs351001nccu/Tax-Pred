# ============================================================================
# 6_variable_importance.R - 計算變數重要性
# ============================================================================

library(data.table)
library(glmnet)
library(ranger)
library(writexl)

# ---- 載入準備好的資料 ----
cat("載入資料...\n")

# 用目前的 h 找資料夾（平行跑才不會亂）
if (!exists("h")) stop("沒有設定 h，請在外面給 h 再跑這支 script。")

results_dir <- file.path("outputs", paste0("h_", h), "results")
if (!dir.exists(results_dir)) stop(paste("找不到", results_dir, "，請先跑對應的 0_setup.R。"))

setup <- readRDS(file.path(results_dir, "prepared_data.rds"))
dat <- setup$dat
target_col <- setup$target_col
time_col <- setup$time_col
lag_vars <- setup$lag_vars
train_len <- setup$train_len
var_exp_cut <- setup$var_exp_cut
elastic_alpha <- setup$elastic_alpha
adapt_gamma <- setup$adapt_gamma
h <- setup$h
output_root <- setup$output_root

cat(sprintf("使用 h = %d 的資料\n", h))

# ---- 參數設定（可調整）----
k_top           <- 8
agg_method      <- "sum"

# 名稱對照設定
map_path        <- meta_path
map_sheet       <- 1
short_max_chars <- 28
short_use_dataset_prefix <- TRUE
dataset_prefix_nchar     <- 6

# ---- glmnet 共用設定 ----
glmnet_common <- list(
  standardize = TRUE,
  maxit = 1e6,
  nlambda = 80,
  lambda.min.ratio = 1e-3
)

# ---- 輔助函數：取前 k 名 ----
topk_names <- function(imp_vec, k = k_top){
  imp_vec <- imp_vec[is.finite(imp_vec)]
  if(length(imp_vec)==0) return(rep(NA_character_, k))
  ord <- order(imp_vec, decreasing=TRUE)
  picks <- names(imp_vec)[ord][seq_len(min(k, length(ord)))]
  out <- rep(NA_character_, k)
  out[seq_along(picks)] <- picks
  out
}

# ===== 建立最後一個 rolling 視窗資料與 direct strategy 的訓練組 =====
cat("\n準備訓練資料...\n")
end_idx   <- nrow(dat)
start_idx <- max(1, end_idx - train_len + 1)
final_tr  <- dat[start_idx:end_idx]

y_all <- final_tr[[target_col]]
X_all <- as.matrix(final_tr[, ..lag_vars])

# 去掉含 NA 的欄、近零變異欄
good_na <- colSums(is.na(X_all)) == 0
X_all   <- X_all[, good_na, drop=FALSE]
if(ncol(X_all) == 0) stop("此視窗內沒有可用的 lag 特徵。")
sd_ok   <- apply(X_all, 2, sd, na.rm=TRUE) > 1e-8
X_all   <- X_all[, sd_ok, drop=FALSE]

# direct strategy：用 (X_t, y_{t+h})
y_lead <- data.table::shift(y_all, type="lead", n=h)
keep   <- !is.na(y_lead)
X_tr2  <- X_all[keep,, drop=FALSE]
y_tr2  <- y_lead[keep]

cat(sprintf("訓練樣本數: %d, 特徵數: %d\n", nrow(X_tr2), ncol(X_tr2)))

# ===== (A) Lasso / Ridge / EN / Adaptive 的重要性（|β|）=====
cat("\n計算 Lasso 家族重要性...\n")

# Lasso
fit_lasso <- tryCatch(
  do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=1), glmnet_common)), 
  error=function(e) NULL
)
imp_lasso <- if(!is.null(fit_lasso)) {
  abs(as.numeric(coef(fit_lasso, s="lambda.1se"))[-1])
} else setNames(numeric(0), character(0))
names(imp_lasso) <- colnames(X_tr2)

# Ridge
fit_ridge <- tryCatch(
  do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=0), glmnet_common)), 
  error=function(e) NULL
)
imp_ridge <- if(!is.null(fit_ridge)) {
  abs(as.numeric(coef(fit_ridge, s="lambda.1se"))[-1])
} else setNames(numeric(0), character(0))
names(imp_ridge) <- colnames(X_tr2)

# Elastic Net
fit_en <- tryCatch(
  do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=elastic_alpha), glmnet_common)), 
  error=function(e) NULL
)
imp_en <- if(!is.null(fit_en)) {
  abs(as.numeric(coef(fit_en, s="lambda.1se"))[-1])
} else setNames(numeric(0), character(0))
names(imp_en) <- colnames(X_tr2)

# Adaptive 權重（Ridge 初估 + 上限裁切）
w_al <- tryCatch({
  fit0 <- do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=0), glmnet_common))
  b0   <- as.numeric(coef(fit0, s="lambda.1se"))[-1]
  pmin(1/(abs(b0)^adapt_gamma + 1e-6), 1e4)
}, error=function(e) rep(1, ncol(X_tr2)))

# Adaptive Lasso
fit_ada_lasso <- tryCatch(
  do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=1, penalty.factor=w_al), glmnet_common)), 
  error=function(e) NULL
)
imp_ada_lasso <- if(!is.null(fit_ada_lasso)) {
  abs(as.numeric(coef(fit_ada_lasso, s="lambda.1se"))[-1])
} else setNames(numeric(0), character(0))
names(imp_ada_lasso) <- colnames(X_tr2)

# Adaptive Elastic Net
fit_ada_en <- tryCatch(
  do.call(glmnet::cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=elastic_alpha, penalty.factor=w_al), glmnet_common)), 
  error=function(e) NULL
)
imp_ada_en <- if(!is.null(fit_ada_en)) {
  abs(as.numeric(coef(fit_ada_en, s="lambda.1se"))[-1])
} else setNames(numeric(0), character(0))
names(imp_ada_en) <- colnames(X_tr2)

# ===== (B) PCR 的重要性（loading × 迴歸係數 的加權和）=====
cat("計算 PCR 重要性...\n")

imp_pcr <- setNames(numeric(0), character(0))
pca <- tryCatch(prcomp(X_tr2, center=TRUE, scale.=TRUE), error=function(e) NULL)
if(!is.null(pca)){
  var_exp <- cumsum(pca$sdev^2)/sum(pca$sdev^2)
  k_pc <- which(var_exp >= var_exp_cut)[1]
  Z_tr <- pca$x[,1:k_pc, drop=FALSE]
  fit_pcr <- tryCatch(lm(y_tr2 ~ Z_tr), error=function(e) NULL)
  if(!is.null(fit_pcr)){
    beta <- coef(fit_pcr)[-1]
    rot  <- pca$rotation[,1:k_pc, drop=FALSE]
    sc   <- as.numeric(abs(rot %*% abs(beta)))
    imp_pcr <- setNames(sc, rownames(rot))
  }
}

# ===== (C) 樹模型重要性：Bagging（proxy）與 Random Forest（Permutation importance）=====
cat("計算樹模型重要性...\n")

p <- ncol(X_tr2)

# Bagging 的 proxy
fit_bag_imp <- tryCatch(
  ranger(y = y_tr2, x = data.frame(X_tr2),
         num.trees = 500, mtry = p, min.node.size = 5,
         importance = "permutation", seed = 123),
  error=function(e) NULL
)
imp_bag <- if(!is.null(fit_bag_imp)) {
  fit_bag_imp$variable.importance
} else setNames(numeric(0), character(0))

# Random Forest
fit_rf_imp <- tryCatch(
  ranger(y = y_tr2, x = data.frame(X_tr2),
         num.trees = 500, mtry = max(1, floor(sqrt(p))), min.node.size = 5,
         importance = "permutation", seed = 123),
  error=function(e) NULL
)
imp_rf <- if(!is.null(fit_rf_imp)) {
  fit_rf_imp$variable.importance
} else setNames(numeric(0), character(0))

# ===== (D)【Lag 層級】組成前 k 名表（代碼名）=====
cat("\n組成 Lag 層級排名表...\n")

imp_list_lag <- list(
  "LASSO"            = imp_lasso,
  "EN"               = imp_en,
  "Adaptive LASSO"   = imp_ada_lasso,
  "Adaptive EN"      = imp_ada_en,
  "Ridge"            = imp_ridge,
  "PCR"              = imp_pcr,
  "Bagging (proxy)"  = imp_bag,
  "Random Forest"    = imp_rf
)

rank_vec <- as.character(seq_len(k_top))
tbl_lag_code <- data.table(排序 = rank_vec)
for (nm in names(imp_list_lag)){
  tbl_lag_code[[nm]] <- topk_names(imp_list_lag[[nm]], k_top)
}

# ===== (E) 載入對照表並映射成短/長名 =====
cat("載入變數名稱對照表...\n")

map_raw <- as.data.table(readxl::read_excel(map_path, sheet = map_sheet))
setnames(map_raw, old=names(map_raw), new=make.names(names(map_raw)))
stopifnot(all(c("var_label","dataset_title","series_name") %in% names(map_raw)))
setorder(map_raw, var_label)
if (any(duplicated(map_raw$var_label))){
  warning("var_label 在對照表中有重複，將取第一筆。")
  map_raw <- map_raw[!duplicated(var_label)]
}

# 分離 lag 後綴
split_lag <- function(x){
  lag <- ifelse(grepl("_L\\d+$", x), sub(".*_L(\\d+)$","L\\1", x), NA_character_)
  base <- ifelse(is.na(lag), x, sub("_L\\d+$","", x))
  list(base=base, lag=lag)
}

# 產生長/短名
make_display_names <- function(base, lag, mapDT,
                               short_max=short_max_chars,
                               use_prefix=short_use_dataset_prefix,
                               prefix_n=dataset_prefix_nchar){
  m <- mapDT[.(base), on=.(var_label)]
  # 長名
  long <- ifelse(!is.na(m$dataset_title) & !is.na(m$series_name),
                 paste0(m$dataset_title, "｜", m$series_name),
                 base)
  long <- ifelse(!is.na(lag), paste0(long, " (", lag, ")"), long)
  
  # 短名
  short_core <- ifelse(!is.na(m$series_name), m$series_name, base)
  if (use_prefix){
    pref <- ifelse(!is.na(m$dataset_title), substr(m$dataset_title, 1L, prefix_n), "")
    short_core <- ifelse(pref!="", paste0(pref, "｜", short_core), short_core)
  }
  short_core <- ifelse(!is.na(lag), paste0(short_core, " (", lag, ")"), short_core)
  short <- ifelse(nchar(short_core) > short_max,
                  paste0(substr(short_core, 1L, short_max-1L), "…"),
                  short_core)
  list(long=long, short=short)
}

map_var_labels <- function(v, mapDT){
  sp <- split_lag(v)
  nm <- make_display_names(sp$base, sp$lag, mapDT)
  data.table(var_label=v, base=sp$base, lag=sp$lag,
             long_name=nm$long, short_name=nm$short)
}

# 套用到 lag 層級表
model_cols <- setdiff(names(tbl_lag_code), "排序")
all_vars <- unique(unlist(lapply(tbl_lag_code[, ..model_cols], unlist), use.names=FALSE))
all_vars <- all_vars[!is.na(all_vars) & all_vars!=""]
dict_lag <- if(length(all_vars)) map_var_labels(all_vars, map_raw) else data.table()

tbl_lag_short <- copy(tbl_lag_code)
tbl_lag_long  <- copy(tbl_lag_code)
for (mc in model_cols){
  m <- merge(data.table(var_label = tbl_lag_code[[mc]]), dict_lag, by="var_label", all.x=TRUE)
  tbl_lag_short[[mc]] <- ifelse(is.na(m$short_name), paste0(m$var_label, " (未對應)"), m$short_name)
  tbl_lag_long[[mc]]  <- ifelse(is.na(m$long_name),  paste0(m$var_label, " (未對應)"), m$long_name)
}

# ===== (F)【彙總到原始變數】把多個 lag 壓回 base 變數並重新排名 =====
cat("彙總到原始變數層級...\n")

aggregate_importance <- function(imp_vec, method = c("sum","max")){
  method <- match.arg(method)
  if(length(imp_vec)==0) return(imp_vec)
  nm <- names(imp_vec)
  base <- sub("_L\\d+$","", nm)
  DT <- data.table(base=base, imp=as.numeric(imp_vec))
  agg <- DT[, .(imp = if (method=="sum") sum(imp, na.rm=TRUE) else max(imp, na.rm=TRUE)), by=base]
  setNames(agg$imp, agg$base)
}

imp_list_base <- lapply(imp_list_lag, aggregate_importance, method = agg_method)

# 組表（代碼名）
tbl_base_code <- data.table(排序 = rank_vec)
for (nm in names(imp_list_base)){
  tbl_base_code[[nm]] <- topk_names(imp_list_base[[nm]], k_top)
}

# 建 base 對照字典
all_bases <- unique(unlist(lapply(tbl_base_code[, ..model_cols], unlist), use.names=FALSE))
all_bases <- all_bases[!is.na(all_bases) & all_bases!=""]
dict_base <- if(length(all_bases)){
  m <- map_raw[.(all_bases), on=.(var_label)]
  data.table(var_label = all_bases,
             short_name = {
               core <- ifelse(!is.na(m$series_name), m$series_name, all_bases)
               if (short_use_dataset_prefix){
                 pref <- ifelse(!is.na(m$dataset_title), substr(m$dataset_title, 1L, dataset_prefix_nchar), "")
                 core <- ifelse(pref!="", paste0(pref, "｜", core), core)
               }
               ifelse(nchar(core) > short_max_chars,
                      paste0(substr(core, 1L, short_max_chars-1L), "…"),
                      core)
             },
             long_name  = ifelse(!is.na(m$dataset_title) & !is.na(m$series_name),
                                 paste0(m$dataset_title, "｜", m$series_name),
                                 all_bases))
} else data.table()

tbl_base_short <- copy(tbl_base_code)
tbl_base_long  <- copy(tbl_base_code)
for (mc in model_cols){
  m <- merge(data.table(var_label = tbl_base_code[[mc]]), dict_base, by="var_label", all.x=TRUE)
  tbl_base_short[[mc]] <- ifelse(is.na(m$short_name), paste0(m$var_label, " (未對應)"), m$short_name)
  tbl_base_long[[mc]]  <- ifelse(is.na(m$long_name),  paste0(m$var_label, " (未對應)"), m$long_name)
}

# ===== (G) 輸出結果 =====
cat("\n輸出結果...\n")

excel_dir <- file.path(output_root, "excel_files")
dir.create(excel_dir, recursive = TRUE, showWarnings = FALSE)

today <- format(Sys.Date(), "%Y%m%d")

# LAG 層級
filename_lag <- sprintf("var_importance_LAG_top%d_h%d_%s.xlsx", k_top, h, today)
writexl::write_xlsx(
  list(
    var_name_dictionary_LAG = dict_lag[, .(var_label, base, lag, short_name, long_name)],
    LAG_top_CODE            = tbl_lag_code,
    LAG_top_SHORT           = tbl_lag_short,
    LAG_top_LONG            = tbl_lag_long
  ),
  path = file.path(excel_dir, filename_lag)
)

# BASE 層級
filename_base <- sprintf("var_importance_BASE_top%d_h%d_%s_%s.xlsx", 
                        k_top, h, toupper(agg_method), today)
writexl::write_xlsx(
  list(
    var_name_dictionary_BASE = dict_base[, .(var_label, short_name, long_name)],
    BASE_top_CODE            = tbl_base_code,
    BASE_top_SHORT           = tbl_base_short,
    BASE_top_LONG            = tbl_base_long
  ),
  path = file.path(excel_dir, filename_base)
)

cat("\n變數重要性分析完成！\n")
cat(sprintf("LAG 層級結果: %s\n", file.path(excel_dir, filename_lag)))
cat(sprintf("BASE 層級結果: %s\n", file.path(excel_dir, filename_base)))

# 預覽結果
cat("\n【LAG 層級】前3名（短名）:\n")
print(head(tbl_lag_short, 3))

cat("\n【BASE 層級】前3名（短名）:\n")
print(head(tbl_base_short, 3))
