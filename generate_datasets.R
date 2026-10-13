# =====================================================================
#  PSY 342 Research Project - simulated dataset generator
# =====================================================================
#
#  What it does
#    Reads each group's data-spec spreadsheet (the "data file with the
#    variables named and explained"), checks it, and generates a simulated
#    dataset that
#      (a) follows the spec exactly: sample size, % female, variable ranges,
#          condition labels, whole numbers on Likert-type scales; and
#      (b) contains the effects YOU set from the group's hypotheses
#          (section 2), at realistic sizes, so the data are usable but not
#          "too perfect".
#
#  How to use
#    1. Put the groups' spec files (.xlsx) in the folder SPEC_DIR.
#    2. In section 2, write each group's effects based on their letter of
#       intent (one list per group).
#    3. Run the whole script:  source("generate_datasets.R")
#    4. For each group, OUT_DIR gets  <file>.csv  (the dataset to send the
#       group; opens in SPSS, Excel, R, JASP...) and  <file>_report.txt
#       (what was read, what was fixed, planned vs. observed effects,
#       descriptives).  _summary.csv lists every group with OK / WARN / ERROR,
#       and _contact_groups.txt lists the spec problems to ask each group about.
#
#  Packages:  install.packages(c("readxl", "haven"))
# =====================================================================

suppressPackageStartupMessages({
  library(readxl)
  library(haven)
})


# ---------------------------------------------------------------------
# 1. Settings
# ---------------------------------------------------------------------

SPEC_DIR <- "specs"     # folder with the groups' .xlsx spec files
OUT_DIR  <- "output"    # results go here

PREPOST_R    <- 0.50  # correlation between a pre and post measure of the same variable
SCALE_SPREAD <- 5     # SD of a scale = (max - min) / SCALE_SPREAD  (1-7 scale -> SD = 1.2)
MIN_PER_CELL <- 20    # warn if fewer participants than this per condition

# "Usable" check: after generating, the script runs the test that matches
# each planned effect.  If TRUE, it redraws (new random seed) until every
# planned effect shows up (p < ALPHA, right direction) and every effect set
# to 0 is non-significant.  The seed used is written in the report.
REQUIRE_HYPOTHESES_DETECTABLE <- TRUE
ALPHA     <- 0.05
MAX_TRIES <- 200
# ...and the observed effect must also land close to what you planned
# (so a planned r = .30 doesn't come out as r = .15 just because p < .05)
TOLERANCE_D <- 0.15   # observed d within +/- 0.15 of planned d
TOLERANCE_R <- 0.10   # observed r within +/- 0.10 of planned r

# Decimals: by default a variable whose range is whole numbers (e.g. 1-7)
# gets whole numbers, otherwise 2 decimals.  Override here by variable
# name, e.g. list(Extraversion = 2) for a scale that is an average of items.
DECIMALS <- list()

# Output format: CSV always.  Set to TRUE to also write an SPSS .sav file.
ALSO_WRITE_SAV <- FALSE


# ---------------------------------------------------------------------
# 2. Effects for each group (fill in from their hypotheses)
# ---------------------------------------------------------------------
#  Effect sizes are on the latent scale (SD = 1):
#    d  0.2 small, 0.5 medium, 0.8 large      r  .10 small, .30 medium, .50 large
#
#  diff_effect(dv, higher, lower, d)
#     Group "higher" scores d SDs above group "lower" on dv.  Works for any
#     grouping variable in the spec (Condition, Gender: "Female"/"Male").
#     If dv is a post measure (e.g. PosEmo_T2, PostPE) this is the
#     Condition x Time interaction: groups start equal at pre-test and
#     differ in change.
#  cor_effect(var1, var2, r)
#     Correlation between two continuous variables.
#  time_effect(pre, post, d)
#     Everyone changes by d SDs from pre to post (main effect of Time).
#  moderation_effect(dv, higher, lower, moderator, d_per_sd)
#     The higher-vs-lower group difference on dv grows by d_per_sd for each
#     SD of the (continuous) moderator.
#
#  Setting d = 0 or r = 0 means "this should NOT be significant".
#
#  For each hypothesis, decide whether it comes
#  out supported (d ~ 0.5 / r ~ .30) or not (d = 0 / r = 0).  Always give a
#  group something (at least one supported effect), but not necessarily
#  everything.  The script warns if a group gets no supported effect.
#  List names are matched against the spec file name (part of the name is
#  enough, e.g. "Group3" matches "smith_Group3_DataSpec.xlsx").

diff_effect <- function(dv, higher, lower, d)
  list(type = "diff", dv = dv, higher = higher, lower = lower, d = d)
cor_effect <- function(var1, var2, r)
  list(type = "cor", var1 = var1, var2 = var2, r = r)
time_effect <- function(pre, post, d)
  list(type = "time", pre = pre, post = post, d = d)
moderation_effect <- function(dv, higher, lower, moderator, d_per_sd)
  list(type = "mod", dv = dv, higher = higher, lower = lower,
       moderator = moderator, d = d_per_sd)

EFFECTS <- list(

  "ExampleDataSet" = list(
    diff_effect("PosEmo_T2", higher = "Pos", lower = "Neg", d = 0.5),
    cor_effect("Extraversion", "PosEmo_T1", r = 0.30)
  )

  # , "Group2" = list(
  #     diff_effect("PostPE", higher = "Positive", lower = "Neutral", d = 0.6),
  #     time_effect("PrePE", "PostPE", d = 0.2)
  #   )
)


# =====================================================================
#  Nothing below needs editing
# =====================================================================

# ---------------------------------------------------------------------
# 3. Small helpers
# ---------------------------------------------------------------------

new_log <- function() {
  e <- new.env()
  e$error <- character(0); e$warn <- character(0); e$fix <- character(0)
  e
}
note <- function(log, type, ...) log[[type]] <- c(log[[type]], paste0(...))

# SPSS-safe names: letters, digits, underscore; must start with a letter
clean_name <- function(x) {
  x <- trimws(as.character(x))
  x <- gsub("[^A-Za-z0-9_]+", "_", x)
  x <- gsub("_+", "_", x)
  x <- gsub("^_|_$", "", x)
  x <- ifelse(grepl("^[A-Za-z]", x), x, paste0("V", x))
  substr(x, 1, 64)
}

# Pull numbers out of cells: "200", "50%", "0.5", "1-7", "1 to 7", "18 - 35"
parse_numbers <- function(cells) {
  out <- numeric(0)
  for (s in cells[!is.na(cells)]) {
    s <- gsub("\u2013|\u2014", "-", s)  # en/em dash -> hyphen
    rng <- "^\\s*(-?[0-9]+(\\.[0-9]+)?)\\s*(-|to|TO|To)\\s*(-?[0-9]+(\\.[0-9]+)?)\\s*$"
    if (grepl(rng, s)) {
      out <- c(out, as.numeric(sub(rng, "\\1", s)), as.numeric(sub(rng, "\\4", s)))
    } else {
      m <- regmatches(s, gregexpr("-?[0-9]*\\.?[0-9]+", s))[[1]]
      out <- c(out, suppressWarnings(as.numeric(m)))
    }
  }
  out[!is.na(out)]
}

# Pre/post pairs among continuous variables:
#   PosEmo_T1 / PosEmo_T2,  PrePE / PostPE,  Mood_Pre / Mood_Post, ...
find_pairs <- function(nms) {
  pre_pat  <- "^(Pre|PRE|pre|Before|before|Baseline|baseline)_?|_?(T1|t1|Time1|time1|Pre|PRE|pre|Before|Baseline)$"
  post_pat <- "^(Post|POST|post|After|after)_?|_?(T2|t2|Time2|time2|Post|POST|post|After)$"
  pre  <- nms[grepl(pre_pat, nms)]
  post <- nms[grepl(post_pat, nms)]
  key_pre  <- sub(pre_pat, "", pre)
  key_post <- sub(post_pat, "", post)
  pairs <- list()
  for (i in seq_along(pre)) {
    m <- post[key_post == key_pre[i]]
    if (length(m) == 1 && m != pre[i]) pairs[[pre[i]]] <- m
  }
  pairs  # named list: pre -> post
}

partner_of <- function(v, pairs) {
  if (v %in% names(pairs)) return(pairs[[v]])
  hit <- names(pairs)[unlist(pairs) == v]
  if (length(hit)) hit[1] else NULL
}

pre_of <- function(v, pairs) {
  hit <- names(pairs)[unlist(pairs) == v]
  if (length(hit)) hit[1] else NULL
}

cohen_d <- function(a, b) {
  sp <- sqrt(((length(a) - 1) * var(a) + (length(b) - 1) * var(b)) /
               (length(a) + length(b) - 2))
  if (!is.finite(sp) || sp == 0) return(NA_real_)
  (mean(a) - mean(b)) / sp
}

seed_from_name <- function(x) {
  v <- utf8ToInt(x)
  (sum(v * seq_along(v)) %% 100000) * 100
}

var_roles <- function(spec) vapply(spec$vars, function(v) v$role, "")
var_names <- function(spec) vapply(spec$vars, function(v) v$name, "")


# ---------------------------------------------------------------------
# 4. Read and check a spec file
# ---------------------------------------------------------------------

read_spec <- function(path) {
  log  <- new_log()
  spec <- list(file = path, N = NA_integer_, female_pct = NA_real_,
               vars = list(), pairs = list())

  sheets <- readxl::excel_sheets(path)
  sheet  <- if ("Data" %in% sheets) "Data" else sheets[1]
  raw <- suppressMessages(readxl::read_excel(path, sheet = sheet, col_names = FALSE,
                                             col_types = "text", .name_repair = "minimal"))
  raw <- as.data.frame(raw, stringsAsFactors = FALSE)
  raw[] <- lapply(raw, function(x) { x <- trimws(x); x[x == ""] <- NA; x })
  if (nrow(raw) > 0) raw <- raw[rowSums(!is.na(raw)) > 0, colSums(!is.na(raw)) > 0, drop = FALSE]
  if (nrow(raw) < 2 || ncol(raw) < 2) {
    note(log, "error", "The spec sheet '", sheet, "' is empty or has no values under the variable names.")
    return(list(spec = spec, log = log))
  }

  header <- unlist(raw[1, ], use.names = FALSE)
  header <- ifelse(is.na(header), paste0("Col", seq_along(header)), header)
  body   <- raw[-1, , drop = FALSE]

  # The template's instruction row ("Range (lower range and then upper range)", ...)
  instr_pat <- "(?i)range|sample size|percent|list cond|lower|upper|no spaces"
  is_instr  <- apply(body, 1, function(r) any(grepl(instr_pat, r[!is.na(r)], perl = TRUE)))

  nms <- clean_name(header)
  if (anyDuplicated(nms)) {
    new <- make.unique(nms, sep = "_")
    note(log, "fix", "Duplicate variable names made unique: ", paste(new[new != nms], collapse = ", "), ".")
    nms <- new
  }
  lnames <- tolower(nms)
  id_col <- which(grepl("^(id|subj|participant|pid|sub$|sub_)", lnames))[1]
  if (is.na(id_col)) id_col <- 1
  gender_col <- which(grepl("gender|^sex", lnames))[1]
  demo_pat <- "race|ethnic|year|class|major|education|religio|marital"

  for (j in seq_along(header)) {
    h  <- header[j]; nm <- nms[j]
    if (nm != h && !grepl("^Col[0-9]+$", h))
      note(log, "fix", "Renamed variable '", h, "' to '", nm, "' (SPSS names can't contain spaces or symbols).")
    if (grepl("^Col[0-9]+$", h))
      note(log, "warn", "Column ", j, " has no variable name; called it '", nm, "'.")

    cells_all <- body[[j]][!is.na(body[[j]])]
    cells     <- body[[j]][!is_instr & !is.na(body[[j]])]
    nums      <- parse_numbers(cells_all)

    # --- ID column: holds the sample size
    if (j == id_col) {
      if (!length(nums)) {
        note(log, "error", "Sample size is missing (expected a number under '", h, "').")
      } else {
        N <- nums[1]
        if (N != round(N) || N < 10) {
          note(log, "error", "Sample size '", N, "' is not usable (must be a whole number of at least 10).")
        } else if (N > 10000) {
          note(log, "error", "Sample size ", N, " is implausibly large for this project.")
        } else {
          if (N > 1000) note(log, "warn", "Sample size ", N, " is very large for a class study.")
          spec$N <- as.integer(N)
        }
      }
      spec$vars[[length(spec$vars) + 1]] <- list(name = nm, label = h, role = "id")
      next
    }

    # --- Gender column: holds % female
    if (!is.na(gender_col) && j == gender_col) {
      if (!length(nums)) {
        note(log, "error", "Percentage female is missing (expected a number under '", h, "').")
        next
      }
      p <- nums[1]
      if (p > 0 && p <= 1 && !any(grepl("%", cells_all))) {
        note(log, "fix", "Read female value ", p, " as a proportion (", p * 100, "% female).")
        p <- p * 100
      }
      if (p < 0 || p > 100) {
        note(log, "error", "Percentage female '", p, "' must be between 0 and 100.")
        next
      }
      if (p == 0 || p == 100) note(log, "warn", "Sample is ", p, "% female, so Gender has no variation.")
      spec$female_pct <- p
      spec$vars[[length(spec$vars) + 1]] <- list(name = nm, label = h, role = "gender",
                                                  levels = c("Female", "Male"))
      next
    }

    # --- Nothing filled in (e.g. Race left blank)
    if (!length(cells) && !length(nums)) {
      note(log, "warn", "Variable '", h, "' has no specification, so it was left out of the dataset.")
      next
    }

    has_text <- any(grepl("[A-Za-z]", gsub("(?i)\\bto\\b", "", cells, perl = TRUE)))

    # --- Categorical: conditions (crossed, balanced) or demographics (random)
    if (has_text) {
      lv <- unlist(strsplit(cells, "\\s*[,;/]\\s*|\\s+(and|&)\\s+"))
      lv <- trimws(lv); lv <- lv[lv != ""]
      lv_clean <- clean_name(lv)
      if (any(lv_clean != lv))
        note(log, "fix", "Cleaned labels in '", nm, "': ",
             paste0(lv[lv_clean != lv], " -> ", lv_clean[lv_clean != lv], collapse = ", "), ".")
      lv_clean <- unique(lv_clean)
      if (length(lv_clean) < 2) {
        note(log, "error", "'", nm, "' lists only one level (", lv_clean, "); a grouping variable needs at least 2.")
        next
      }
      role <- if (grepl(demo_pat, lnames[j])) "demo" else "cat"
      if (role == "demo")
        note(log, "warn", "'", nm, "' treated as a demographic variable; its categories were drawn with equal probability.")
      if (length(lv_clean) > 6)
        note(log, "warn", "'", nm, "' has ", length(lv_clean), " levels, which is a lot for this project.")
      spec$vars[[length(spec$vars) + 1]] <- list(name = nm, label = h, role = role, levels = lv_clean)
      next
    }

    # --- Numeric range
    if (length(nums) < 2) {
      note(log, "error", "'", nm, "' needs a lower and an upper value (found only ", length(nums), ").")
      next
    }
    if (length(nums) > 2)
      note(log, "warn", "'", nm, "' has more than two numbers; used the first two (", nums[1], ", ", nums[2], ") as the range.")
    lo <- min(nums[1:2]); hi <- max(nums[1:2])
    if (nums[1] > nums[2]) note(log, "fix", "Range of '", nm, "' was entered upper-first; swapped to ", lo, "-", hi, ".")
    if (lo == hi) {
      note(log, "error", "'", nm, "' has the same lower and upper value (", lo, "), so it can't vary.")
      next
    }
    is_age <- grepl("^age", lnames[j])
    if (is_age) {
      if (ceiling(lo) > floor(hi)) {
        note(log, "error", "Age range ", lo, "-", hi, " contains no whole-number age.")
        next
      }
      if (lo < 16 || hi > 100) note(log, "warn", "Age range ", lo, "-", hi, " looks unusual for a participant sample.")
    }
    dec <- if (!is.null(DECIMALS[[nm]])) DECIMALS[[nm]] else if (all(c(lo, hi) == round(c(lo, hi)))) 0 else 2
    spec$vars[[length(spec$vars) + 1]] <- list(name = nm, label = h, role = if (is_age) "age" else "num",
                                                lo = lo, hi = hi, decimals = dec)
  }

  # --- Whole-spec checks
  roles <- var_roles(spec)
  if (is.na(spec$N) && !length(log$error)) note(log, "error", "Could not find the sample size.")
  if (!any(roles == "num")) note(log, "error", "No continuous (rating-scale) variables with a range, so there is nothing to analyze.")
  if (!any(roles == "cat")) note(log, "warn", "No condition variable found (only demographics/continuous variables).")
  if (!is.na(spec$N) && any(roles == "cat")) {
    cells_n <- prod(vapply(spec$vars[roles == "cat"], function(v) length(v$levels), 0))
    if (spec$N / cells_n < MIN_PER_CELL)
      note(log, "warn", "Only about ", floor(spec$N / cells_n), " participants per condition (", cells_n,
           " conditions); fewer than ", MIN_PER_CELL, " per condition makes effects hard to detect.")
  }
  spec$pairs <- find_pairs(var_names(spec)[roles == "num"])
  list(spec = spec, log = log)
}


# ---------------------------------------------------------------------
# 5. Check the effects you wrote against the spec
# ---------------------------------------------------------------------

resolve_effects <- function(effects, spec, log) {
  roles <- var_roles(spec); nms <- var_names(spec)
  cont  <- nms[roles == "num"]
  grp   <- spec$vars[roles %in% c("cat", "gender", "demo")]
  names(grp) <- vapply(grp, function(v) v$name, "")
  out <- list()

  group_n <- function(by, level) {
    v <- grp[[by]]
    if (v$role == "gender") {
      nf <- round(spec$N * spec$female_pct / 100)
      return(if (level == "Female") nf else spec$N - nf)
    }
    spec$N / length(v$levels)
  }

  for (e in effects) {
    vf <- intersect(c("dv", "var1", "var2", "pre", "post", "moderator"), names(e))
    for (f in vf) e[[f]] <- clean_name(e[[f]])
    for (f in intersect(c("higher", "lower"), names(e))) e[[f]] <- clean_name(e[[f]])
    desc <- paste0(e$type, "(", paste(unlist(e[c(vf, intersect(c("higher", "lower"), names(e)))]), collapse = ", "), ")")

    miss <- setdiff(unlist(e[vf]), cont)
    if (length(miss)) {
      note(log, "error", "Effect ", desc, ": '", paste(miss, collapse = "', '"),
           "' is not a continuous variable in this spec (available: ", paste(cont, collapse = ", "), ").")
      next
    }
    if (e$type %in% c("diff", "mod")) {
      hit <- names(grp)[vapply(grp, function(v) all(c(e$higher, e$lower) %in% v$levels), logical(1))]
      if (!length(hit)) {
        lv <- vapply(grp, function(v) paste0(v$name, ": ", paste(v$levels, collapse = "/")), "")
        note(log, "error", "Effect ", desc, ": no grouping variable has both '", e$higher, "' and '", e$lower,
             "' (available: ", paste(lv, collapse = "; "), ").")
        next
      }
      e$by <- hit[1]
    }
    if (e$type == "cor") {
      if (abs(e$r) >= 1) { note(log, "error", "Effect ", desc, ": r must be between -1 and 1."); next }
      if (abs(e$r) > 0.6) note(log, "warn", "Effect ", desc, ": r = ", e$r, " is unusually strong for psychology data.")
      if (e$r != 0 && !is.na(spec$N)) {
        pw <- pnorm(atanh(abs(e$r)) * sqrt(spec$N - 3) - qnorm(1 - ALPHA / 2))
        if (pw < 0.5) note(log, "warn", "Effect ", desc, ": with N = ", spec$N, " an r of ", e$r,
                          " would be detected only ~", round(100 * pw), "% of the time; if redrawn until significant, the observed r will be inflated.")
      }
    } else {
      if (abs(e$d) > 1.2) note(log, "warn", "Effect ", desc, ": d = ", e$d,
                               " is very large (typical psychology effects are d = 0.2-0.8).")
      if (e$type == "diff" && e$d != 0 && !is.na(spec$N)) {
        n <- min(group_n(e$by, e$higher), group_n(e$by, e$lower))
        pw <- power.t.test(n = n, delta = abs(e$d), sig.level = ALPHA)$power
        if (pw < 0.5) note(log, "warn", "Effect ", desc, ": with ~", floor(n), " per group a d of ", e$d,
                          " would be detected only ~", round(100 * pw), "% of the time; if redrawn until significant, the observed d will be inflated.")
      }
    }
    out[[length(out) + 1]] <- e
  }
  out
}


# ---------------------------------------------------------------------
# 6. Simulate
# ---------------------------------------------------------------------

# Correlation matrix of the continuous variables (same for every redraw)
build_R <- function(spec, eff, log) {
  cont <- var_names(spec)[var_roles(spec) == "num"]
  R <- diag(length(cont)); dimnames(R) <- list(cont, cont)
  set <- R == 1
  for (pre in names(spec$pairs)) {
    post <- spec$pairs[[pre]]
    R[pre, post] <- R[post, pre] <- PREPOST_R
  }
  for (e in eff) if (e$type == "cor") {
    R[e$var1, e$var2] <- R[e$var2, e$var1] <- e$r
    set[e$var1, e$var2] <- set[e$var2, e$var1] <- TRUE
  }
  # Spill-over: a variable correlated with a pre-test also correlates
  # (more weakly) with the matching post-test, unless you set that pair yourself
  for (e in eff) if (e$type == "cor") {
    for (pr in list(c(e$var1, e$var2), c(e$var2, e$var1))) {
      partner <- partner_of(pr[2], spec$pairs)
      if (!is.null(partner) && partner != pr[1] && !set[pr[1], partner] && R[pr[1], partner] == 0)
        R[pr[1], partner] <- R[partner, pr[1]] <- e$r * PREPOST_R
    }
  }
  k <- 0
  off <- row(R) != col(R)
  while (min(eigen(R, symmetric = TRUE, only.values = TRUE)$values) < 1e-6 && k < 100) {
    R[off] <- 0.95 * R[off]; k <- k + 1
  }
  if (k > 0) note(log, "warn", "The correlations you set can't all hold at once; they were shrunk by ",
                  round(100 * (1 - 0.95^k)), "% to make them consistent.")
  R
}

draw_age <- function(N, lo, hi) {
  lo <- ceiling(lo); hi <- floor(hi)
  ctr <- lo + 0.25 * (hi - lo); s <- max((hi - lo) / 4, 0.5)  # skewed toward the young end, like a student sample
  a <- round(rnorm(N, ctr, s))
  bad <- a < lo | a > hi
  while (any(bad)) { a[bad] <- round(rnorm(sum(bad), ctr, s)); bad <- a < lo | a > hi }
  a
}

simulate_once <- function(spec, eff, R, seed) {
  set.seed(seed)
  N <- spec$N
  vars <- spec$vars; names(vars) <- var_names(spec)
  roles <- var_roles(spec)
  out <- list()

  # Conditions: fully crossed, balanced, random order
  cat_vars <- vars[roles == "cat"]
  if (length(cat_vars)) {
    grid <- expand.grid(lapply(cat_vars, function(v) v$levels), stringsAsFactors = FALSE)
    cell <- sample(rep_len(sample(nrow(grid)), N))
    for (nm in names(cat_vars)) out[[nm]] <- grid[[nm]][cell]
  }
  for (v in vars) {
    if (v$role == "id") out[[v$name]] <- seq_len(N)
    if (v$role == "gender") {
      nf <- round(N * spec$female_pct / 100)
      out[[v$name]] <- sample(rep(c("Female", "Male"), c(nf, N - nf)))
    }
    if (v$role == "demo") out[[v$name]] <- sample(v$levels, N, replace = TRUE)
    if (v$role == "age")  out[[v$name]] <- draw_age(N, v$lo, v$hi)
  }

  # Continuous variables: correlated latent z-scores + effects, then mapped to the scale
  cont <- names(vars)[roles == "num"]
  Z <- matrix(rnorm(N * length(cont)), N) %*% chol(R[cont, cont, drop = FALSE])
  colnames(Z) <- cont
  for (e in eff) {
    if (e$type == "diff") {
      g <- out[[e$by]]
      Z[g == e$higher, e$dv] <- Z[g == e$higher, e$dv] + e$d / 2
      Z[g == e$lower,  e$dv] <- Z[g == e$lower,  e$dv] - e$d / 2
    }
    if (e$type == "time") Z[, e$post] <- Z[, e$post] + e$d
    if (e$type == "mod") {
      g <- out[[e$by]]
      s <- ifelse(g == e$higher, 0.5, ifelse(g == e$lower, -0.5, 0))
      Z[, e$dv] <- Z[, e$dv] + e$d * s * Z[, e$moderator]
    }
  }
  for (v in vars[cont]) {
    sdv <- (v$hi - v$lo) / SCALE_SPREAD
    x <- round((v$lo + v$hi) / 2 + sdv * Z[, v$name], v$decimals)
    out[[v$name]] <- pmin(pmax(x, v$lo), v$hi)
  }
  data.frame(out[names(vars)], stringsAsFactors = FALSE, check.names = FALSE)
}


# ---------------------------------------------------------------------
# 7. Test that the planned effects are in the data
# ---------------------------------------------------------------------

# For a post measure with a matching pre measure, test change scores:
# for two groups this is the same test as the Condition x Time interaction.
outcome <- function(dat, dv, pairs) {
  pre <- pre_of(dv, pairs)
  if (is.null(pre)) list(y = dat[[dv]], label = dv)
  else list(y = dat[[dv]] - dat[[pre]], label = paste0("change ", pre, "->", dv))
}

test_effects <- function(dat, eff, pairs) {
  if (!length(eff)) return(data.frame(effect = character(0), stat = character(0), planned = numeric(0),
                                      observed = numeric(0), p = numeric(0), pass = logical(0)))
  rows <- lapply(eff, function(e) {
    if (e$type == "diff") {
      o <- outcome(dat, e$dv, pairs); g <- dat[[e$by]]
      a <- o$y[g == e$higher]; b <- o$y[g == e$lower]
      p <- t.test(a, b)$p.value; est <- cohen_d(a, b); stat <- "d"
      lab <- paste0(o$label, ": ", e$higher, " vs ", e$lower,
                    if (!is.null(pre_of(e$dv, pairs))) paste0(" (", e$by, " x Time)") else "")
    } else if (e$type == "cor") {
      ct <- cor.test(dat[[e$var1]], dat[[e$var2]])
      p <- ct$p.value; est <- unname(ct$estimate); stat <- "r"
      lab <- paste0(e$var1, " with ", e$var2)
    } else if (e$type == "time") {
      x1 <- dat[[e$pre]]; x2 <- dat[[e$post]]
      p <- t.test(x2, x1, paired = TRUE)$p.value
      est <- mean(x2 - x1) / sqrt((var(x1) + var(x2)) / 2); stat <- "d"
      lab <- paste0(e$pre, " -> ", e$post, " (Time)")
    } else {
      o <- outcome(dat, e$dv, pairs)
      keep <- dat[[e$by]] %in% c(e$higher, e$lower)
      g <- as.numeric(dat[[e$by]][keep] == e$higher)
      m <- as.numeric(scale(dat[[e$moderator]][keep]))
      fit <- summary(lm(o$y[keep] ~ g * m))
      p <- fit$coefficients["g:m", "Pr(>|t|)"]
      est <- fit$coefficients["g:m", "Estimate"] / fit$sigma; stat <- "d/SD"
      lab <- paste0(o$label, ": ", e$higher, " vs ", e$lower, " x ", e$moderator)
    }
    target <- if (e$type == "cor") e$r else e$d
    tol  <- if (e$type == "cor") TOLERANCE_R else TOLERANCE_D
    pass <- if (target == 0) isTRUE(p >= ALPHA) else
            isTRUE(p < ALPHA && sign(est) == sign(target) && abs(est - target) <= tol)
    data.frame(effect = lab, stat = stat, planned = target, observed = round(est, 2),
               p = signif(p, 3), pass = pass, stringsAsFactors = FALSE)
  })
  do.call(rbind, rows)
}


# ---------------------------------------------------------------------
# 8. Write .sav / .csv / report
# ---------------------------------------------------------------------

write_data <- function(dat, spec, base) {
  sav <- dat
  for (v in spec$vars) {
    if (v$role %in% c("gender", "cat", "demo")) {
      codes <- match(sav[[v$name]], v$levels)
      sav[[v$name]] <- haven::labelled(codes, labels = setNames(seq_along(v$levels), v$levels), label = v$label)
    } else {
      attr(sav[[v$name]], "label") <- v$label
    }
  }
  write.csv(dat, paste0(base, ".csv"), row.names = FALSE)
  if (ALSO_WRITE_SAV) haven::write_sav(sav, paste0(base, ".sav"))
}

describe <- function(dat, spec) {
  roles <- var_roles(spec); nms <- var_names(spec)
  lines <- character(0)
  for (v in spec$vars[roles %in% c("gender", "cat", "demo")]) {
    tb <- table(factor(dat[[v$name]], levels = v$levels))
    lines <- c(lines, paste0("  ", v$name, ": ", paste0(names(tb), " = ", tb, collapse = ", ")))
  }
  by <- nms[roles == "cat"][1]
  for (v in nms[roles %in% c("age", "num")]) {
    x <- dat[[v]]
    s <- sprintf("  %-14s M = %.2f, SD = %.2f, min %s, max %s", v, mean(x), sd(x), min(x), max(x))
    if (!is.na(by)) {
      lv <- spec$vars[[which(nms == by)]]$levels
      s <- paste0(s, paste0(sprintf(" | %s: M = %.2f, SD = %.2f", lv,
                                    vapply(lv, function(l) mean(x[dat[[by]] == l]), 0),
                                    vapply(lv, function(l) sd(x[dat[[by]] == l]), 0)), collapse = ""))
    }
    lines <- c(lines, s)
  }
  lines
}

realism_checks <- function(dat, spec, log) {
  for (v in spec$vars[var_roles(spec) == "num"]) {
    x <- dat[[v$name]]
    if (sd(x) == 0) note(log, "warn", "'", v$name, "' has no variation.")
    for (b in c(v$lo, v$hi)) {
      share <- mean(x == b)
      if (share > 0.25) note(log, "warn", round(100 * share), "% of '", v$name, "' sit at ", b,
                             " (floor/ceiling); the effect may be too large for this scale.")
    }
  }
}


# ---------------------------------------------------------------------
# 9. Run one spec, then all specs
# ---------------------------------------------------------------------

effects_for <- function(file) {
  hit <- names(EFFECTS)[vapply(names(EFFECTS), function(k) grepl(k, file, fixed = TRUE), logical(1))]
  if (length(hit)) EFFECTS[[hit[1]]] else NULL
}

generate_one <- function(path) {
  file <- basename(path)
  base <- file.path(OUT_DIR, tools::file_path_sans_ext(file))
  parsed <- read_spec(path); spec <- parsed$spec; log <- parsed$log
  # Spec problems to ask the group to fix
  log$ask <- c(log$error, grep("no specification|has no variable name|more than two numbers|looks unusual",
                               log$warn, value = TRUE))

  effects <- effects_for(file)
  if (is.null(effects)) {
    note(log, "warn", "No effects set for this file in EFFECTS, so the data contain no built-in effects (pure noise).")
    effects <- list()
  }
  eff <- if (!length(log$error)) resolve_effects(effects, spec, log) else list()
  if (!length(log$error) && !any(vapply(eff, function(e) (if (e$type == "cor") e$r else e$d) != 0, logical(1))))
    note(log, "warn", "This group gets no supported effect. Every group should get at least one: set at least one effect above 0.")

  tests <- NULL; dat <- NULL; seed <- NA; tries <- 0
  if (!length(log$error)) {
    R <- build_R(spec, eff, log)
    base_seed <- seed_from_name(file)
    best <- NULL
    for (k in seq_len(MAX_TRIES)) {
      sim <- simulate_once(spec, eff, R, base_seed + k)
      tst <- test_effects(sim, eff, spec$pairs)
      score <- sum(tst$pass)
      if (is.null(best) || score > best$score) best <- list(dat = sim, tests = tst, score = score, seed = base_seed + k)
      tries <- k
      if (!REQUIRE_HYPOTHESES_DETECTABLE || score == nrow(tst)) break
    }
    dat <- best$dat; tests <- best$tests; seed <- best$seed
    if (nrow(tests) && !all(tests$pass))
      note(log, "warn", "After ", tries, " tries not every planned effect came out as intended (see table); ",
           "consider a larger d / r, or a larger N.")
    realism_checks(dat, spec, log)
    dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
    write_data(dat, spec, base)
  }

  status <- if (length(log$error)) "ERROR" else if (length(log$warn)) "WARN" else "OK"

  txt <- c(paste0("Spec file: ", file), paste0("Status:    ", status),
           paste0("Generated: ", format(Sys.time(), "%Y-%m-%d %H:%M")), "")
  if (length(log$error)) txt <- c(txt, "ERRORS (no dataset was generated - fix the spec or EFFECTS):", paste0("  - ", log$error), "")
  if (length(log$ask))   txt <- c(txt, "ASK THE GROUP about (see _contact_groups.txt):", paste0("  - ", log$ask), "")
  if (length(log$fix))   txt <- c(txt, "Automatically fixed:", paste0("  - ", log$fix), "")
  if (length(log$warn))  txt <- c(txt, "Warnings:", paste0("  - ", log$warn), "")
  if (!is.null(dat)) {
    txt <- c(txt, sprintf("N = %d   seed = %s   tries = %d", spec$N, seed, tries), "")
    if (length(spec$pairs)) txt <- c(txt, paste0("Pre/post pairs (r = ", PREPOST_R, "): ",
                                            paste0(names(spec$pairs), " -> ", unlist(spec$pairs), collapse = ", ")), "")
    if (nrow(tests)) {
      txt <- c(txt, "Planned vs. observed effects:")
      txt <- c(txt, sprintf("  [%s] %-48s planned %s = %5.2f   observed = %5.2f   p = %s",
                            ifelse(tests$pass, "ok", "!!"), tests$effect, tests$stat,
                            tests$planned, tests$observed, format(tests$p)), "")
    }
    txt <- c(txt, "Descriptives:", describe(dat, spec))
  }
  dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
  writeLines(txt, paste0(base, "_report.txt"))

  data.frame(file = file, status = status, N = spec$N,
             effects_ok = if (is.null(tests)) NA else paste0(sum(tests$pass), "/", nrow(tests)),
             tries = tries, errors = length(log$error), warnings = length(log$warn),
             first_problem = c(log$error, log$warn, "")[1],
             ask_group = paste(log$ask, collapse = " || "), stringsAsFactors = FALSE)
}

run_all <- function() {
  files <- list.files(SPEC_DIR, pattern = "\\.xlsx?$", full.names = TRUE, ignore.case = TRUE)
  files <- files[!grepl("^~\\$", basename(files))]  # skip Excel lock files
  if (!length(files)) stop("No .xlsx files found in '", SPEC_DIR, "'.")
  summary <- do.call(rbind, lapply(files, function(f) {
    message("Processing ", basename(f), " ...")
    tryCatch(generate_one(f), error = function(err)
      data.frame(file = basename(f), status = "ERROR", N = NA, effects_ok = NA, tries = 0,
                 errors = 1, warnings = 0, first_problem = paste("Script error:", conditionMessage(err)),
                 ask_group = "",
                 stringsAsFactors = FALSE))
  }))
  dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
  write.csv(summary, file.path(OUT_DIR, "_summary.csv"), row.names = FALSE)

  # One list of everything to ask the groups about, ready to paste into a message
  ask <- summary[summary$ask_group != "", ]
  contact <- if (!nrow(ask)) "No spec problems - nothing to ask the groups about." else
    unlist(lapply(seq_len(nrow(ask)), function(i) c(
      paste0("=== ", ask$file[i], if (ask$status[i] == "ERROR") "  (no dataset generated yet)" else "  (dataset generated; please confirm)"),
      "Hi! Before I can send your dataset, could you clarify the following in your data file?",
      paste0("  - ", strsplit(ask$ask_group[i], " || ", fixed = TRUE)[[1]]),
      "")))
  writeLines(contact, file.path(OUT_DIR, "_contact_groups.txt"))
  print(summary[, c("file", "status", "N", "effects_ok", "tries", "first_problem")], row.names = FALSE)
  invisible(summary)
}

run_all()
