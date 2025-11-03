# ============================================================================
# 0_setup.R - 資料準備與前處理
# ============================================================================

# ---- 安裝/載入套件 ----
pkgs <- c("data.table","readxl","lubridate")
ins <- pkgs[!pkgs %in% installed.packages()[,"Package"]]
if(length(ins)) install.packages(ins, repos="https://cloud.r-project.org")
invisible(lapply(pkgs, library, character.only=TRUE))

local_path <- getwd()

# ---- 參數設定 ----
excel_path    <- file.path(local_path, "data", "wide_tab.xlsx")
sheet_name    <- 1
time_col      <- "ym"
target_col    <- "A100101010_1"
if (!exists("h")) {
  h <- 1   # 外面沒給 h 的話，就用預設值 1
}
cat(">>> 現在的 h =", h, "\n")
train_len     <- 60
p_max         <- 6
var_exp_cut   <- 0.95
elastic_alpha <- 0.5
adapt_gamma   <- 1.0

meta_path <- file.path(local_path, "data", "meta_tab.xlsx")

# ---- 建立 h 專屬輸出目錄 ----
output_root <- file.path(local_path, "outputs", sprintf("h_%d", h))
dir.create(output_root, recursive = TRUE, showWarnings = FALSE)

results_dir <- file.path(output_root, "results")
dir.create(results_dir, showWarnings = FALSE)

cat(sprintf("\n=== 輸出目錄設定 ===\n"))
cat(sprintf("h = %d\n", h))
cat(sprintf("輸出根目錄: %s\n\n", output_root))

# ---- 讀檔 ----
cat("讀取資料...\n")
raw <- as.data.table(readxl::read_excel(excel_path, sheet = sheet_name))
setnames(raw, old=names(raw), new=make.names(names(raw)))
time_col   <- make.names(time_col)
target_col <- make.names(target_col)

stopifnot(time_col %in% names(raw), target_col %in% names(raw))

# ---- 解析時間欄 (YYYY-Mm 格式) ----
parse_ym_M <- function(v){
  v  <- as.character(v)
  yr <- as.integer(sub("^(\\d{4})-M.*$", "\\1", v))
  mm <- as.integer(sub("^\\d{4}-M(\\d{1,2})$", "\\1", v))
  if (any(is.na(yr)) || any(is.na(mm))) {
    stop("時間欄位必須長得像 '2025-M1' 或 '2025-M11'")
  }
  as.Date(sprintf("%04d-%02d-01", yr, mm))
}

raw[, (time_col) := parse_ym_M(get(time_col))]
data.table::setorderv(raw, cols = time_col, order = 1L, na.last = TRUE)

# ---- 儲存原始稅收值 ----
cat("\n儲存原始稅收值...\n")
raw[, paste0(target_col, "_original") := get(target_col)]

# ---- 轉換為 YoY ----
cat("轉換目標變數為 YoY...\n")
raw[, (target_col) := {
  original <- get(paste0(target_col, "_original"))
  lag12 <- shift(original, n=12, type="lag")
  (original - lag12) / lag12
}]

cat(sprintf("目標變數 %s 已轉為 YoY 成長率\n", target_col))
cat("前 15 個值:\n")
print(head(raw[, c(time_col, target_col, paste0(target_col, "_original")), with=FALSE], 15))

# ---- 變數轉換（使其平穩）----
cat("\n變數轉換...\n")
meta_data <- readxl::read_excel(meta_path)
tcode_mapping <- setNames(meta_data$tcode, meta_data$var_label)

transform_variable <- function(x, tcode) {
  if(tcode == 1) return(x)
  else if(tcode == 2) return(c(NA, diff(x)))
  else if(tcode == 3) {
    temp <- c(NA, diff(x))
    return(c(NA, diff(temp, na.rm=FALSE)))
  }
  else if(tcode == 4) return(log(x))
  else if(tcode == 5) return(c(NA, diff(log(x))))
  else if(tcode == 6) {
    temp <- c(NA, diff(log(x)))
    return(c(NA, diff(temp, na.rm=FALSE)))
  }
  else if(tcode == 7) {
    pct <- x / c(NA, x[-length(x)]) - 1
    return(c(NA, diff(pct)))
  }
  else return(x)
}

transform_vars <- setdiff(names(raw), c(time_col, target_col, paste0(target_col, "_original")))

for(v in transform_vars) {
  tcode <- tcode_mapping[v]
  if(!is.na(tcode)) {
    raw[[v]] <- transform_variable(raw[[v]], tcode)
  }
}

cat("轉換完成\n")

# ---- 計算 PCA 因子 ----
cat("\n計算 PCA 因子...\n")
numeric_vars <- setdiff(names(raw), c(time_col, target_col, paste0(target_col, "_original")))
numeric_data <- as.data.frame(raw[, ..numeric_vars])

complete_rows <- complete.cases(numeric_data)
clean_data <- numeric_data[complete_rows, ]
cat("用於 PCA 的樣本數:", sum(complete_rows), "/", nrow(raw), "\n")

scaled_data <- scale(clean_data)
pca_result <- prcomp(scaled_data, center=FALSE, scale.=FALSE)
factors <- pca_result$x[, 1:4]

var_explained <- summary(pca_result)$importance[2, 1:4]
cat("前4個因子解釋變異:\n")
cat(sprintf("  Factor1: %.1f%%\n", var_explained[1]*100))
cat(sprintf("  Factor2: %.1f%%\n", var_explained[2]*100))
cat(sprintf("  Factor3: %.1f%%\n", var_explained[3]*100))
cat(sprintf("  Factor4: %.1f%%\n", var_explained[4]*100))

factor_df <- as.data.frame(matrix(NA, nrow=nrow(raw), ncol=4))
names(factor_df) <- paste0("Factor", 1:4)
factor_df[complete_rows, ] <- factors
raw <- cbind(raw, factor_df)

# ---- 建立 lag 特徵 ----
cat("\n建立 lag 特徵...\n")
num_vars <- setdiff(names(raw), c(time_col, paste0(target_col, "_original")))
make_lags <- function(dt, vars, pmax){
  dt <- copy(dt)
  for (v in vars){
    for (L in 1:pmax){
      dt[, paste0(v,"_L",L) := shift(get(v), n=L, type="lag")]
    }
  }
  dt
}
dat <- make_lags(raw, vars=num_vars, pmax=p_max)
lag_vars <- grep("(_L[0-9]+)$", names(dat), value=TRUE)

cat(sprintf("建立完成：%d 個 lag 特徵\n", length(lag_vars)))

# ---- 儲存處理好的資料 ----
cat("\n儲存處理後資料...\n")

saveRDS(list(
  dat = dat,
  target_col = target_col,
  target_col_original = paste0(target_col, "_original"),
  time_col = time_col,
  lag_vars = lag_vars,
  train_len = train_len,
  h = h,
  var_exp_cut = var_exp_cut,
  elastic_alpha = elastic_alpha,
  adapt_gamma = adapt_gamma,
  output_root = output_root,
  results_dir = results_dir
), file.path(results_dir, "prepared_data.rds"))

cat(sprintf("資料已儲存至 %s\n", file.path(results_dir, "prepared_data.rds")))
cat("維度:", nrow(dat), "列 ×", ncol(dat), "欄\n")
