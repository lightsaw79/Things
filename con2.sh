#!/bin/ksh
# controller_backlog_with_state.ksh
# Usage: controller_backlog_with_state.ksh <BOX_NAME>
#
# REQUIRED env:
#   AUTOSYS_RESET_DIR          # dir with autorep.sh and sendevent
#   MODEL_BATCH_LOGFILES_DIR   # writable dir for logs/state
#   BATCH_TABLE_FILE           # ETL/SAM table (whitespace columns)
#
# OPTIONAL env:
#   DAILY_START_HHMM=14:00     # daily box start time (HH:MM, 24h)
#   CUTOFF_BUFFER_MIN=30       # stop starting new backlog if within N minutes of next start
#   RUN_SLEEP_SEC=30           # poll interval while RU/AC
#   RETRY_SLEEP_SEC=10         # poll interval for other/unknown
#
# Notes:
# - Pending and history are newline-separated lists of dates (e.g., 10-MAR-2025)

box_name="$1"

# --- guards ---
[ -n "${AUTOSYS_RESET_DIR:-}" ] || exit 0
[ -n "${MODEL_BATCH_LOGFILES_DIR:-}" ] || exit 0
[ -s "${BATCH_TABLE_FILE:-}" ] || exit 0

: "${DAILY_START_HHMM:=14:00}"
: "${CUTOFF_BUFFER_MIN:=30}"
: "${RUN_SLEEP_SEC:=30}"
: "${RETRY_SLEEP_SEC:=10}"

logdir="$MODEL_BATCH_LOGFILES_DIR"
STATUS_LOG="${logdir}/${box_name}_get_status.log"
PENDING_FILE="${logdir}/${box_name}_bdates.lst"
HISTORY_FILE="${logdir}/${box_name}_bdates_history.lst"

# Ensure state files exist
[ -f "$PENDING_FILE" ] || : > "$PENDING_FILE"
[ -f "$HISTORY_FILE" ] || : > "$HISTORY_FILE"

# --- helper: minutes until next daily start ---
_mins_until_next_start() {
  shh="${DAILY_START_HHMM%:*}"; smm="${DAILY_START_HHMM#*:}"
  nh="$(date +%H)"; nm="$(date +%M)"
  now=$((10#$nh*60 + 10#$nm))
  tgt=$((10#$shh*60 + 10#$smm))
  if [ $now -lt $tgt ]; then echo $((tgt - now)); else echo $((24*60 - now + tgt)); fi
}

# --- always write a FRESH get-status log, then parse ST/Ex code (RU/AC/SU/FA/IN/OH/OI/TE) ---
_refresh_status_log() {
  _tmp="${STATUS_LOG}.$$"
  "${AUTOSYS_RESET_DIR}/autorep.sh" -q "$box_name" > "${_tmp}" 2>&1
  echo "TS=$(date +'%Y-%m-%d %H:%M:%S')" >> "${_tmp}"
  mv "${_tmp}" "${STATUS_LOG}"
}
_get_status_code() {
  _refresh_status_log
  awk -v box="$box_name" '
    /ST\/Ex/ { st=index($0,"ST/Ex"); next }
    st>0 && $0 !~ /^[- ]*$/ {
      name=substr($0,1,st-1); gsub(/^ +| +$/,"",name)
      if (name==box) {
        c=substr($0,st,2); gsub(/[^A-Za-z]/,"",c)
        print toupper(c); exit
      }
    }' "${STATUS_LOG}"
}

# --- inline force-start (your exact steps) for BACKLOG only ---
_force_start_now() {
  "${AUTOSYS_RESET_DIR}/sendevent" -p 1 -E CHANGE_STATUS -s INACTIVE -J "$box_name" \
    > "${logdir}/${box_name}_inactive.log" 2>&1
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "$box_name" \
    > "${logdir}/${box_name}.log" 2>&1
}

# --- wait until SU (used for backlog runs). If failure-ish, re-kick. ---
_wait_until_su() {
  while : ; do
    S="$(_get_status_code)"
    case "$S" in
      RU|AC)  sleep "$RUN_SLEEP_SEC" ;;
      SU)     return 0 ;;
      FA|TE|IN|OH|OI|'') _force_start_now; sleep "$RETRY_SLEEP_SEC" ;;
      *)      _force_start_now; sleep "$RETRY_SLEEP_SEC" ;;
    esac
  done
}

# --- Build candidate backlog from table (1 date per line), ignoring ETL STARTED ---
_candidates_from_table() {
  awk '
    function is_date(x){
      return (x ~ /^[0-3]?[0-9]-[A-Za-z]{3}-[0-9]{2,4}$/ || x ~ /^[0-9]{4}-[01][0-9]-[0-3][0-9]$/)
    }
    /_SAM_BATCH/ && $0 ~ /COMPLETED/ { sam[$2]=1; next }
    /_ETL_BATCH/ && $0 ~ /COMPLETED/ {
      d=""; for(i=NF;i>=1;i--) if(is_date($i)){ d=$i; break }
      if(d!=""){ bd[$2]=d; order[++n]=$2 }
    }
    END {
      for(i=1;i<=n;i++){ id=order[i]; if(!(id in sam)) print bd[id] }
    }
  ' "$BATCH_TABLE_FILE"
}

# --- Merge: keep existing pending; append any new candidates not in pending/history ---
_merge_pending_with_new() {
  existing="$(cat "$PENDING_FILE")"
  # build a quick membership check for history
  while IFS= read -r d; do
    [ -n "$d" ] || continue
    # skip if already done
    if grep -qx "$d" "$HISTORY_FILE" 2>/dev/null; then
      continue
    fi
    # skip if already pending
    if [ -n "$existing" ] && echo "$existing" | grep -qx "$d"; then
      continue
    fi
    existing="${existing}${existing:+
}$d"
  done <<EOF
$(_candidates_from_table)
EOF
  # write back atomically
  _tmp="${PENDING_FILE}.$$"
  printf "%s\n" "$existing" | sed '/^$/d' > "$_tmp"
  mv "$_tmp" "$PENDING_FILE"
}

# --- pop head from pending (returns via global _HEAD); rewrites pending file ---
_pop_head() {
  _HEAD="$(sed -n '1p' "$PENDING_FILE")"
  _tmp="${PENDING_FILE}.$$"
  sed '1d' "$PENDING_FILE" > "$_tmp"
  mv "$_tmp" "$PENDING_FILE"
}

# --- append to history (dedupe-safe append) ---
_append_history() {
  d="$1"
  [ -n "$d" ] || return 0
  # only append if not already present
  if ! grep -qx "$d" "$HISTORY_FILE" 2>/dev/null; then
    printf "%s\n" "$d" >> "$HISTORY_FILE"
  fi
}

# ========================== MAIN FLOW ==========================

# Step A: Merge fresh candidates into pending (keeps leftovers + adds new)
_merge_pending_with_new

# Step B: PHASE 1 — Wait for the MAIN-DAY run to finish (no force-start here)
while : ; do
  # Respect cutoff even while waiting for main
  mins="$(_mins_until_next_start)"
  if [ "$mins" -le "$CUTOFF_BUFFER_MIN" ]; then
    exit 0
  fi

  S0="$(_get_status_code)"
  case "$S0" in
    RU|AC)  sleep "$RUN_SLEEP_SEC" ;;      # main run executing
    SU)     break ;;                        # main-day done → proceed to backlog
    FA|TE|IN|OH|OI|'') sleep "$RETRY_SLEEP_SEC" ;;  # do not interfere with main-day
    *)      sleep "$RETRY_SLEEP_SEC" ;;
  esac
done

# Step C: PHASE 2 — Drain BACKLOG dates one-by-one, honoring cutoff
while : ; do
  # stop if close to next start
  mins="$(_mins_until_next_start)"
  if [ "$mins" -le "$CUTOFF_BUFFER_MIN" ]; then
    exit 0
  fi

  # load next pending date
  _pop_head
  [ -n "$_HEAD" ] || break   # nothing pending → done

  # kick backlog run for this date and wait for SU
  _force_start_now
  _wait_until_su

  # move this date into history (so it won't reappear tomorrow)
  _append_history "$_HEAD"
done

# Done (either cleared backlog or hit cutoff earlier)
exit 0