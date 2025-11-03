# ============================================================================
# 3_models_ml.R - 機器學習模型 (Bagging Tree / Random Forest)
# ============================================================================

library(data.table)
library(ipred)
library(ranger)

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
h <- setup$h
output_root <- setup$output_root

cat(sprintf("使用 h = %d 的資料\n", h))

# ---- 指標函數 ----
rmse <- function(e) sqrt(mean(e^2, na.rm=TRUE))
mae  <- function(e) mean(abs(e), na.rm=TRUE)
madm <- function(e) median(abs(e), na.rm=TRUE)

# ---- Rolling 預測 ----
cat("\n開始 Rolling 預測...\n")
n <- nrow(dat)
idx_train_end <- seq(from = train_len, to = n - h, by = 1)
if (!length(idx_train_end)) stop("資料太短，請降低 train_len 或確認資料長度。")

collect_ml <- list()
pb <- txtProgressBar(min = 0, max = length(idx_train_end), style = 3)

for (i in seq_along(idx_train_end)){
  te <- idx_train_end[i]
  train_idx <- (te - train_len + 1):te
  test_idx  <- te + h

  tr     <- dat[train_idx]
  te_row <- dat[test_idx]

  y_tr <- tr[[target_col]]
  y_te <- te_row[[target_col]]

  # 準備特徵矩陣
  X_tr <- as.matrix(tr[, ..lag_vars])
  X_te <- as.matrix(te_row[, ..lag_vars])
  good <- colSums(is.na(X_tr)) == 0
  X_tr <- X_tr[, good, drop=FALSE]
  X_te <- X_te[, good, drop=FALSE]
  
  # 過濾近零變異
  if(ncol(X_tr) > 0){
    sd_ok <- apply(X_tr, 2, sd, na.rm=TRUE) > 1e-8
    X_tr  <- X_tr[, sd_ok, drop=FALSE]
    X_te  <- X_te[, sd_ok, drop=FALSE]
  }

  if (ncol(X_tr) >= 1){
    tr_df <- data.frame(y = y_tr, X_tr)
    te_df <- data.frame(X_te)

    # Bagging
    fit_bag <- tryCatch(
      ipred::bagging(y ~ ., data=tr_df, coob=TRUE, nbagg=100),
      error=function(e) NULL
    )
    yhat_bag <- if(!is.null(fit_bag)) as.numeric(predict(fit_bag, newdata=te_df)) else NA_real_

    # Random Forest
    fit_rf <- tryCatch(
      ranger::ranger(y ~ ., data=tr_df,
                     num.trees=500,
                     mtry=max(1, floor(sqrt(ncol(X_tr)))),
                     min.node.size=5),
      error=function(e) NULL
    )
    yhat_rf <- if(!is.null(fit_rf)) as.numeric(predict(fit_rf, data=te_df)$predictions) else NA_real_

    collect_ml[[length(collect_ml)+1]] <- rbind(
      data.table(date = te_row[[time_col]], model="BaggingTree",  y=y_te, yhat=yhat_bag),
      data.table(date = te_row[[time_col]], model="RandomForest", y=y_te, yhat=yhat_rf)
    )
  }
  
  setTxtProgressBar(pb, i)
}
close(pb)

# ---- 彙整結果 ----
roll_ml <- rbindlist(collect_ml, fill=TRUE)
roll_ml <- roll_ml[!is.na(yhat)]

cat("\n計算指標...\n")
score_ml <- roll_ml[, .(
  RMSE = rmse(y - yhat),
  MAE  = mae(y - yhat),
  MAD  = madm(y - yhat)
), by=.(model)]

print(score_ml)

# ---- 儲存結果 ----
cat("\n儲存結果...\n")
saveRDS(list(
  roll_ml = roll_ml,
  score_ml = score_ml
), file.path(results_dir, "roll_ml.rds"))

cat(sprintf("機器學習模型完成！結果已儲存至 %s\n", 
            file.path(results_dir, "roll_ml.rds")))
