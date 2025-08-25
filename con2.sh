#!/bin/ksh
# controller_drain_from_table_codes_with_cutoff.ksh
# Usage: controller_drain_from_table_codes_with_cutoff.ksh <BOX_NAME>
#
# REQUIRED env:
#   AUTOSYS_RESET_DIR          # dir with autorep.sh and sendevent
#   MODEL_BATCH_LOGFILES_DIR   # writable dir for logs
#   BATCH_TABLE_FILE           # ETL/SAM table (whitespace columns)
#
# OPTIONAL env (defaults shown):
#   DAILY_START_HHMM=14:00     # next day start time of the box (HH:MM, 24h)
#   CUTOFF_BUFFER_MIN=30       # stop starting new dates if within this many minutes of next start

box_name="$1"

# --- guards ---
[ -n "${AUTOSYS_RESET_DIR:-}" ] || exit 0
[ -n "${MODEL_BATCH_LOGFILES_DIR:-}" ] || exit 0
[ -s "${BATCH_TABLE_FILE:-}" ] || exit 0

logdir="$MODEL_BATCH_LOGFILES_DIR"
STATUS_LOG="${logdir}/${box_name}_get_status.log"
: "${DAILY_START_HHMM:=14:00}"
: "${CUTOFF_BUFFER_MIN:=30}"

# --- helpers: minutes until next daily start (no GNU date needed) ---
_mins_until_next_start() {
  shh="${DAILY_START_HHMM%:*}"; smm="${DAILY_START_HHMM#*:}"
  nh="$(date +%H)"; nm="$(date +%M)"
  now=$((10#$nh*60 + 10#$nm))
  tgt=$((10#$shh*60 + 10#$smm))
  if [ $now -lt $tgt ]; then
    echo $((tgt - now))
  else
    echo $((24*60 - now + tgt))
  fi
}

# --- build BDATES: ETL COMPLETED with NO SAM COMPLETED (match by App ID) ---
BDATES="$(
  awk '
    function is_date(x){
      return (x ~ /^[0-3]?[0-9]-[A-Za-z]{3}-[0-9]{2,4}$/ || x ~ /^[0-9]{4}-[01][0-9]-[0-3][0-9]$/)
    }
    /_SAM_BATCH/ && $0 ~ /COMPLETED/ { sam[$2]=1; next }
    /_ETL_BATCH/ && $0 ~ /COMPLETED/ {
      d=""; for(i=NF;i>=1;i--) if(is_date($i)){ d=$i; break }
      if(d!=""){ bd[$2]=d; order[++n]=$2 }
    }
    END{
      first=1
      for(i=1;i<=n;i++){
        id=order[i]
        if(!(id in sam)){ if(!first) printf " "; printf "%s", bd[id]; first=0 }
      }
    }
  ' "$BATCH_TABLE_FILE"
)"
[ -n "$BDATES" ] || exit 0

# --- always write a FRESH get-status log, then read ST/Ex code for this box ---
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
      if(name==box){
        c=substr($0,st,2); gsub(/[^A-Za-z]/,"",c)
        print toupper(c); exit
      }
    }' "${STATUS_LOG}"
}

# --- inline force-start (your exact steps) ---
_force_start_now() {
  "${AUTOSYS_RESET_DIR}/sendevent" -p 1 -E CHANGE_STATUS -s INACTIVE -J "$box_name" \
    > "${logdir}/${box_name}_inactive.log" 2>&1
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "$box_name" \
    > "${logdir}/${box_name}.log" 2>&1
}

# --- wait until SU, refreshing logs each poll (RU/AC waits 30s) ---
_wait_until_su() {
  while : ; do
    S="$(_get_status_code)"
    case "$S" in
      RU|AC)  sleep 30 ;;           # running → wait here
      SU)     return 0 ;;           # success → continue
      FA|TE|IN|OH|OI|'')            # failed/iced/hold/inactive/blank → recover
              _force_start_now; sleep 10 ;;
      *)      _force_start_now; sleep 10 ;;
    esac
  done
}

# --- if currently running, wait-in-place until not RU/AC (never exit to next hour) ---
while : ; do
  S0="$(_get_status_code)"
  case "$S0" in RU|AC) sleep 30 ;; *) break ;; esac
done

# --- drain ALL pending BDATES in THIS run, but honor cutoff before starting each new date ---
for _d in $BDATES ; do
  # stop if we are too close to the next daily start (do NOT start a new date)
  mins=$(_mins_until_next_start)
  if [ "$mins" -le "$CUTOFF_BUFFER_MIN" ]; then
    # leave remaining dates for the next window (they’ll still show as ETL-without-SAM)
    exit 0
  fi

  # start next date now
  _force_start_now

  # wait until this run completes successfully (may bounce RU/AC/…)
  _wait_until_su
done

# all queued dates finished in this window
exit 0