# ============================================================================
# 2_models_lasso.R - Lasso 家族 (Lasso/Ridge/EN/Adaptive/PCR)
# ============================================================================

library(data.table)
library(glmnet)

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
var_exp_cut <- setup$var_exp_cut
elastic_alpha <- setup$elastic_alpha
adapt_gamma <- setup$adapt_gamma
output_root <- setup$output_root

cat(sprintf("使用 h = %d 的資料\n", h))

# ---- 指標函數 ----
rmse <- function(e) sqrt(mean(e^2, na.rm=TRUE))
mae  <- function(e) mean(abs(e), na.rm=TRUE)
madm <- function(e) median(abs(e), na.rm=TRUE)

# ---- glmnet 公共參數 ----
glmnet_common <- list(
  standardize = TRUE,
  maxit = 1e6,
  nlambda = 80,
  lambda.min.ratio = 1e-3
)

# ---- Rolling 預測 ----
cat("\n開始 Rolling 預測...\n")
n <- nrow(dat)
idx_train_end <- seq(from = train_len, to = n - h, by = 1)
if (!length(idx_train_end)) stop("資料太短，請降低 train_len 或確認資料長度。")

collect_lasso <- list()
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
    # Lasso
    fit_lasso <- tryCatch(
      do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=1), glmnet_common)),
      error=function(e) NULL
    )
    yhat_lasso <- if(!is.null(fit_lasso)) as.numeric(predict(fit_lasso, X_te, s="lambda.1se")) else NA_real_

    # Ridge
    fit_ridge <- tryCatch(
      do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=0), glmnet_common)),
      error=function(e) NULL
    )
    yhat_ridge <- if(!is.null(fit_ridge)) as.numeric(predict(fit_ridge, X_te, s="lambda.1se")) else NA_real_
    
    # Elastic Net
    fit_en <- tryCatch(
      do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=elastic_alpha), glmnet_common)),
      error=function(e) NULL
    )
    yhat_en <- if(!is.null(fit_en)) as.numeric(predict(fit_en, X_te, s="lambda.1se")) else NA_real_
    
    # Adaptive 權重（以 Ridge 初估）
    w_al <- tryCatch({
      fit0 <- do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=0), glmnet_common))
      b0   <- as.numeric(coef(fit0, s="lambda.1se"))[-1]
      w    <- 1/(abs(b0)^adapt_gamma + 1e-6)
      pmin(w, 1e4)
    }, error=function(e) rep(1, ncol(X_tr)))
    
    # Adaptive Lasso
    fit_ada_lasso <- tryCatch(
      do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=1, penalty.factor=w_al), glmnet_common)),
      error=function(e) NULL
    )
    yhat_ada_lasso <- if(!is.null(fit_ada_lasso)) as.numeric(predict(fit_ada_lasso, X_te, s="lambda.1se")) else NA_real_
    
    # Adaptive Elastic Net
    fit_ada_en <- tryCatch(
      do.call(cv.glmnet, c(list(x=X_tr, y=y_tr, alpha=elastic_alpha, penalty.factor=w_al), glmnet_common)),
      error=function(e) NULL
    )
    yhat_ada_en <- if(!is.null(fit_ada_en)) as.numeric(predict(fit_ada_en, X_te, s="lambda.1se")) else NA_real_

    # PCR
    pca <- tryCatch(prcomp(X_tr, center=TRUE, scale.=TRUE), error=function(e) NULL)
    yhat_pcr <- NA_real_
    if(!is.null(pca)){
      var_exp <- cumsum(pca$sdev^2)/sum(pca$sdev^2)
      k <- which(var_exp >= var_exp_cut)[1]
      Z_tr <- pca$x[,1:k,drop=FALSE]
      Z_te <- scale(X_te, center=pca$center, scale=pca$scale) %*% pca$rotation[,1:k,drop=FALSE]
      fit_pcr <- tryCatch(lm(y_tr ~ Z_tr), error=function(e) NULL)
      if(!is.null(fit_pcr)) yhat_pcr <- as.numeric(cbind(1, Z_te) %*% coef(fit_pcr))
    }

    collect_lasso[[length(collect_lasso)+1]] <- data.table(
      date = te_row[[time_col]],
      model = c("Lasso","Ridge","ElasticNet","AdaLasso","AdaElasticNet","PCR"),
      y     = y_te,
      yhat  = c(yhat_lasso,yhat_ridge,yhat_en,yhat_ada_lasso,yhat_ada_en,yhat_pcr)
    )
  }
  
  setTxtProgressBar(pb, i)
}
close(pb)

# ---- 彙整結果 ----
roll_lasso <- rbindlist(collect_lasso, fill=TRUE)
roll_lasso <- roll_lasso[!is.na(yhat)]

cat("\n計算指標...\n")
score_lasso <- roll_lasso[, .(
  RMSE = rmse(y - yhat),
  MAE  = mae(y - yhat),
  MAD  = madm(y - yhat)
), by=.(model)]

print(score_lasso)

# ---- 儲存結果 ----
cat("\n儲存結果...\n")
saveRDS(list(
  roll_lasso = roll_lasso,
  score_lasso = score_lasso
), file.path(results_dir, "roll_lasso.rds"))

cat(sprintf("Lasso 家族模型完成！結果已儲存至 %s\n", 
            file.path(results_dir, "roll_lasso.rds")))
