# ============================================================================
# 1_models_ts.R - 時間序列模型 (RandomWalk / AR / ARMA)
# ============================================================================

library(data.table)
library(forecast)

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

collect_ts <- list()
pb <- txtProgressBar(min = 0, max = length(idx_train_end), style = 3)

for (i in seq_along(idx_train_end)){
  te <- idx_train_end[i]
  train_idx <- (te - train_len + 1):te
  test_idx  <- te + h

  tr     <- dat[train_idx]
  te_row <- dat[test_idx]

  y_tr <- tr[[target_col]]
  y_te <- te_row[[target_col]]

  # Random Walk
  yhat_rw <- tail(y_tr, 1)

  # AR
  fit_ar <- tryCatch(
    forecast::auto.arima(y_tr, d=0, max.q=0, seasonal=FALSE,
                        stepwise=TRUE, approximation=TRUE), 
    error=function(e) NULL
  )
  yhat_ar <- tryCatch(
    as.numeric(forecast::forecast(fit_ar, h=h)$mean[h]), 
    error=function(e) yhat_rw
  )

  # ARMA
  fit_arma <- tryCatch(
    forecast::auto.arima(y_tr, d=0, seasonal=FALSE,
                        stepwise=TRUE, approximation=TRUE), 
    error=function(e) NULL
  )
  yhat_arma <- tryCatch(
    as.numeric(forecast::forecast(fit_arma, h=h)$mean[h]), 
    error=function(e) yhat_rw
  )

  collect_ts[[length(collect_ts)+1]] <- data.table(
    date = te_row[[time_col]],
    model = c("RandomWalk","AR","ARMA"),
    y     = y_te,
    yhat  = c(yhat_rw, yhat_ar, yhat_arma)
  )
  
  setTxtProgressBar(pb, i)
}
close(pb)

# ---- 彙整結果 ----
roll_ts <- rbindlist(collect_ts, fill=TRUE)
roll_ts <- roll_ts[!is.na(yhat)]

cat("\n計算指標...\n")
score_ts <- roll_ts[, .(
  RMSE = rmse(y - yhat),
  MAE  = mae(y - yhat),
  MAD  = madm(y - yhat)
), by=.(model)]

print(score_ts)

# ---- 儲存結果 ----
cat("\n儲存結果...\n")
saveRDS(list(
  roll_ts = roll_ts,
  score_ts = score_ts
), file.path(results_dir, "roll_ts.rds"))

cat(sprintf("時間序列模型完成！結果已儲存至 %s\n", 
            file.path(results_dir, "roll_ts.rds")))
