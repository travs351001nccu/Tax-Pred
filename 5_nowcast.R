# ============================================================================
# 5_nowcast.R - 未來多期預測（Nowcast）
# ============================================================================

library(data.table)
library(lubridate)
library(forecast)
library(glmnet)
library(ipred)
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
target_col_original <- setup$target_col_original
time_col <- setup$time_col
lag_vars <- setup$lag_vars
train_len <- setup$train_len
var_exp_cut <- setup$var_exp_cut
elastic_alpha <- setup$elastic_alpha
adapt_gamma <- setup$adapt_gamma
h <- setup$h
output_root <- setup$output_root

cat(sprintf("使用 h = %d 的資料\n", h))

# ---- 參數：一次預測未來 H_out 個月 ----
H_out <- 12

# ---- glmnet 公共設定 ----
glmnet_common <- list(
  standardize=TRUE, 
  maxit=1e6, 
  nlambda=80, 
  lambda.min.ratio=1e-3
)

# ---- 準備最後一個 rolling 訓練窗 ----
cat("\n準備最後一個訓練窗...\n")
last_time <- max(dat[[time_col]])
end_idx   <- nrow(dat)
start_idx <- max(1, end_idx - train_len + 1)
final_tr  <- dat[start_idx:end_idx]

# 目標向量與特徵矩陣
y_all   <- final_tr[[target_col]]
X_all   <- as.matrix(final_tr[, ..lag_vars])
good    <- colSums(is.na(X_all)) == 0
X_all   <- X_all[, good, drop=FALSE]

# 預測當下可用的特徵（來自全資料最後一列的 lag）
last_row <- tail(dat, 1)
x_origin <- as.matrix(last_row[, ..lag_vars])[, good, drop=FALSE]

# 預測的日期（t+1 ... t+H_out）
future_dates <- sapply(1:H_out, function(hh) last_time %m+% months(hh))

# ===== (A) 時間序列：一次多步 =====
cat("\n時間序列模型預測...\n")
next_rw_path <- rep(tail(y_all,1), H_out)

fit_ar_all <- tryCatch(
  forecast::auto.arima(y_all, d=0, max.q=0, seasonal=FALSE), 
  error=function(e) NULL
)
next_ar_path <- tryCatch(
  as.numeric(forecast::forecast(fit_ar_all, h=H_out)$mean), 
  error=function(e) next_rw_path
)

fit_arma_all <- tryCatch(
  forecast::auto.arima(y_all, d=0, seasonal=FALSE), 
  error=function(e) NULL
)
next_arma_path <- tryCatch(
  as.numeric(forecast::forecast(fit_arma_all, h=H_out)$mean), 
  error=function(e) next_rw_path
)

extra_ts <- data.table(
  date  = as.Date(future_dates, origin = "1970-01-01"),
  horizon = 1:H_out,
  model = rep(c("RandomWalk","AR","ARMA"), each = H_out),
  yhat  = c(next_rw_path, next_ar_path, next_arma_path)
)

# ===== (B) Lasso 家族 + PCR：direct strategy =====
cat("Lasso 家族模型預測...\n")

# 建 H_out 組 training pairs
train_list <- lapply(1:H_out, function(hh){
  y_h  <- shift(y_all, type="lead", n=hh)
  keep <- !is.na(y_h)
  list(y = y_h[keep], X = X_all[keep,, drop=FALSE], h = hh)
})

res_lasso <- list()
pb <- txtProgressBar(min = 0, max = H_out, style = 3)

for (idx in seq_along(train_list)){
  item <- train_list[[idx]]
  y_tr2 <- item$y
  X_tr2 <- item$X
  hh    <- item$h

  # 過濾近零變異
  sd_ok <- apply(X_tr2, 2, sd, na.rm=TRUE) > 1e-8
  X_tr2 <- X_tr2[, sd_ok, drop=FALSE]
  x0    <- x_origin[, sd_ok, drop=FALSE]

  if (ncol(X_tr2) < 1){
    res_lasso[[length(res_lasso)+1]] <- data.table(
      date = future_dates[hh], horizon=hh,
      model = c("Lasso","Ridge","ElasticNet","AdaLasso","AdaElasticNet","PCR"),
      yhat  = NA_real_
    )
    setTxtProgressBar(pb, idx)
    next
  }

  # Lasso / Ridge / EN
  fit_lasso <- tryCatch(
    do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=1), glmnet_common)), 
    error=function(e) NULL
  )
  yhat_lasso <- if(!is.null(fit_lasso)) as.numeric(predict(fit_lasso, x0, s="lambda.1se")) else NA_real_

  fit_ridge <- tryCatch(
    do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=0), glmnet_common)), 
    error=function(e) NULL
  )
  yhat_ridge <- if(!is.null(fit_ridge)) as.numeric(predict(fit_ridge, x0, s="lambda.1se")) else NA_real_

  fit_en <- tryCatch(
    do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=elastic_alpha), glmnet_common)), 
    error=function(e) NULL
  )
  yhat_en <- if(!is.null(fit_en)) as.numeric(predict(fit_en, x0, s="lambda.1se")) else NA_real_

  # Adaptive 權重
  w_al <- tryCatch({
    fit0 <- do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=0), glmnet_common))
    b0   <- as.numeric(coef(fit0, s="lambda.1se"))[-1]
    w    <- 1/(abs(b0)^adapt_gamma + 1e-6)
    pmin(w, 1e4)
  }, error=function(e) rep(1, ncol(X_tr2)))

  fit_ada_lasso <- tryCatch(
    do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=1, penalty.factor=w_al), glmnet_common)), 
    error=function(e) NULL
  )
  yhat_ada_lasso <- if(!is.null(fit_ada_lasso)) as.numeric(predict(fit_ada_lasso, x0, s="lambda.1se")) else NA_real_

  fit_ada_en <- tryCatch(
    do.call(cv.glmnet, c(list(x=X_tr2, y=y_tr2, alpha=elastic_alpha, penalty.factor=w_al), glmnet_common)), 
    error=function(e) NULL
  )
  yhat_ada_en <- if(!is.null(fit_ada_en)) as.numeric(predict(fit_ada_en, x0, s="lambda.1se")) else NA_real_

  # PCR
  yhat_pcr <- NA_real_
  pca <- tryCatch(prcomp(X_tr2, center=TRUE, scale.=TRUE), error=function(e) NULL)
  if(!is.null(pca)){
    var_exp <- cumsum(pca$sdev^2)/sum(pca$sdev^2)
    k <- which(var_exp >= var_exp_cut)[1]
    Z_tr <- pca$x[,1:k,drop=FALSE]
    fit_pcr <- tryCatch(lm(y_tr2 ~ Z_tr), error=function(e) NULL)
    if(!is.null(fit_pcr)){
      z0 <- scale(x0, center=pca$center, scale=pca$scale) %*% pca$rotation[,1:k,drop=FALSE]
      yhat_pcr <- as.numeric(cbind(1, z0) %*% coef(fit_pcr))
    }
  }

  res_lasso[[length(res_lasso)+1]] <- data.table(
    date = future_dates[hh], horizon = hh,
    model = c("Lasso","Ridge","ElasticNet","AdaLasso","AdaElasticNet","PCR"),
    yhat  = c(yhat_lasso,yhat_ridge,yhat_en,yhat_ada_lasso,yhat_ada_en,yhat_pcr)
  )
  
  setTxtProgressBar(pb, idx)
}
close(pb)

extra_lasso <- rbindlist(res_lasso, fill=TRUE)

# ===== (C) 非線性：Bagging Tree + Random Forest (direct strategy) =====
cat("\n機器學習模型預測...\n")

res_ml <- list()
pb <- txtProgressBar(min = 0, max = H_out, style = 3)

for (idx in seq_along(train_list)){
  item <- train_list[[idx]]
  y_tr2 <- item$y
  X_tr2 <- item$X
  hh    <- item$h

  # 過濾近零變異
  sd_ok <- apply(X_tr2, 2, sd, na.rm=TRUE) > 1e-8
  X_tr2 <- X_tr2[, sd_ok, drop=FALSE]
  x0    <- x_origin[, sd_ok, drop=FALSE]

  if (ncol(X_tr2) < 1){
    res_ml[[length(res_ml)+1]] <- data.table(
      date=future_dates[hh], horizon=hh,
      model=c("BaggingTree","RandomForest"),
      yhat=c(NA_real_, NA_real_)
    )
    setTxtProgressBar(pb, idx)
    next
  }

  tr_df <- data.frame(y=y_tr2, X_tr2)
  te_df <- data.frame(x0)

  # Bagging
  fit_bag <- tryCatch(
    ipred::bagging(y ~ ., data=tr_df, coob=FALSE, nbagg=100), 
    error=function(e) NULL
  )
  yhat_bag <- if(!is.null(fit_bag)) as.numeric(predict(fit_bag, newdata=te_df)) else NA_real_

  # Random Forest
  fit_rf <- tryCatch(
    ranger::ranger(y ~ ., data=tr_df, num.trees=500,
                   mtry=max(1, floor(sqrt(ncol(X_tr2)))), min.node.size=5),
    error=function(e) NULL
  )
  yhat_rf <- if(!is.null(fit_rf)) as.numeric(predict(fit_rf, data=te_df)$predictions) else NA_real_

  res_ml[[length(res_ml)+1]] <- data.table(
    date=future_dates[hh], horizon=hh,
    model=c("BaggingTree","RandomForest"),
    yhat=c(yhat_bag,yhat_rf)
  )
  
  setTxtProgressBar(pb, idx)
}
close(pb)

extra_ml <- rbindlist(res_ml, fill=TRUE)

# ===== (D) 整併輸出 =====
cat("\n整併結果...\n")
extra_all <- rbindlist(list(extra_ts, extra_lasso, extra_ml), fill=TRUE)
setorderv(extra_all, c("horizon","model"))

# ===== (E) 將 YoY 預測轉回實際稅收值 =====
cat("將 YoY 預測轉回實際稅收值...\n")

# 取得最近 12 個月的原始稅收值
last_12_original <- tail(dat[[target_col_original]], 12)

extra_all[, actual_forecast := {
  yoy_pred <- yhat
  lag12_index <- ((horizon - 1) %% 12) + 1
  base_value <- last_12_original[lag12_index]
  base_value * (1 + yoy_pred)
}]

extra_all[, base_value_t12 := {
  lag12_index <- ((horizon - 1) %% 12) + 1
  last_12_original[lag12_index]
}]

cat("\n還原後的預測結果（前 24 筆）:\n")
print(head(extra_all[, .(date, horizon, model, yhat_yoy=yhat, 
                         base_value_t12, actual_forecast)], 24))

# ===== (F) 輸出 =====
cat("\n輸出結果...\n")
excel_dir <- file.path(output_root, "excel_files")
dir.create(excel_dir, recursive = TRUE, showWarnings = FALSE)

today <- format(Sys.Date(), "%Y%m%d")
filename <- sprintf("nowcast_H%d_h%d_%s.xlsx", H_out, h, today)

writexl::write_xlsx(
  extra_all[, .(date, horizon, model, 
                yhat_yoy = yhat,
                base_value_t12,
                actual_forecast)],
  path = file.path(excel_dir, filename)
)

cat(sprintf("\n已輸出包含實際稅收預測值的檔案:\n"))
cat(sprintf("  %s\n", file.path(excel_dir, filename)))

cat("\nNowcast 完成！\n")
