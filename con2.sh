#!/bin/ksh
# controller_backlog_with_state.ksh
box_name="$1"
[ -n "${AUTOSYS_RESET_DIR:-}" ] || exit 0
[ -n "${MODEL_BATCH_LOGFILES_DIR:-}" ] || exit 0
: "${STATE_DIR:=/apps/samd/actimize/package_utilities/common/bin}"
: "${DAILY_START_HHMM:=14:00}"
: "${CUTOFF_BUFFER_MIN:=30}"
: "${RUN_SLEEP_SEC:=30}"
: "${RETRY_SLEEP_SEC:=10}"
logdir="$MODEL_BATCH_LOGFILES_DIR"
STATUS_LOG="${logdir}/${box_name}_get_status.log"
STATE_FILE="${STATE_DIR}/${box_name}_bdates.lst"
HISTORY_FILE="${STATE_DIR}/${box_name}_bdates_history.lst"
[ -f "$STATE_FILE" ] || : > "$STATE_FILE"
[ -f "$HISTORY_FILE" ] || : > "$HISTORY_FILE"
TFILE="${BATCH_TABLE_FILE:-}"
if [ -z "$TFILE" ]; then
  TFILE=$(ls -1t "${STATE_DIR}"/*[Bb][Aa][Tt][Cc][Hh]*.{txt,csv,log} 2>/dev/null | head -1)
fi
[ -s "$TFILE" ] || exit 0

mins_until_next_start() {
  shh="${DAILY_START_HHMM%:*}"; smm="${DAILY_START_HHMM#*:}"
  nh="$(date +%H)"; nm="$(date +%M)"
  now=$((10#$nh*60 + 10#$nm))
  tgt=$((10#$shh*60 + 10#$smm))
  if [ $now -lt $tgt ]; then echo $((tgt - now)); else echo $((24*60 - now + tgt)); fi
}

refresh_status_log() {
  _tmp="${STATUS_LOG}.$$"
  "${AUTOSYS_RESET_DIR}/autorep.sh" -q "$box_name" > "${_tmp}" 2>&1
  echo "TS=$(date +'%Y-%m-%d %H:%M:%S')" >> "${_tmp}"
  mv "${_tmp}" "${STATUS_LOG}"
}
get_status_code() {
  refresh_status_log
  tr -d '\r' < "$STATUS_LOG" | awk -v box="$box_name" '
    BEGIN{ st=0; code_first="" }
    /ST\/Ex/ { st=index($0,"ST/Ex"); next }
    st>0 && $0 !~ /^[-= ]*$/ {
      name=substr($0,1,st-1); gsub(/^ +| +$/,"",name)
      c=substr($0,st,2); gsub(/[^A-Za-z]/,"",c); c=toupper(c)
      if (toupper(name)==toupper(box)) { print c; exit }
      if (code_first=="") code_first=c
    }
    END{ if (code_first!="") print code_first }
  '
}


force_start_now() {
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "$box_name" \
    > "${logdir}/${box_name}.log" 2>&1
}

wait_until_su() {
  while : ; do
    S="$(get_status_code)"
    case "$S" in
      RU|AC)  sleep "$RUN_SLEEP_SEC" ;;
      SU)     return 0 ;;
      FA|TE|IN|OH|OI|'') force_start_now; sleep "$RETRY_SLEEP_SEC" ;;
      *)      force_start_now; sleep "$RETRY_SLEEP_SEC" ;;
    esac
  done
}

CANDS="$(
  awk '
    BEGIN{ IGNORECASE=1 }
    function is_date(x){
      return (x ~ /^[0-3]?[0-9]-[A-Za-z]{3}-[0-9]{2,4}$/ || x ~ /^[0-9]{4}-[01][0-9]-[0-3][0-9]$/)
    }
    /_SAM_BATCH/ && $0 ~ /COMPLETED/ { sam[$2]=1; next }
    /_ETL_BATCH/ && $0 ~ /COMPLETED/ {
      d=""; for (i=NF; i>=1; i--) if (is_date($i)) { d=$i; break }
      if (d!="") { bd[$2]=d; order[++n]=$2 }
    }
    END{
      first=1
      for (i=1;i<=n;i++){
        id=order[i]
        if (!(id in sam)) { if(!first) printf " "; printf "%s", bd[id]; first=0 }
      }
    }
  ' "$TFILE"
)"

BDATES="$(cat "$STATE_FILE")"
merged="$( { printf '%s\n' $BDATES; printf '%s\n' $CANDS; } | awk 'NF && !seen[$0]++' )"
if [ -s "$HISTORY_FILE" ]; then
  merged="$( printf '%s\n' $merged | grep -vx -f "$HISTORY_FILE" || true )"
fi
BDATES="$( printf '%s\n' $merged | xargs )"
echo "$BDATES" > "$STATE_FILE"

while : ; do
  mins="$(mins_until_next_start)"
  if [ "$mins" -le "$CUTOFF_BUFFER_MIN" ]; then exit 0; fi
  S0="$(get_status_code)"
  case "$S0" in
    RU|AC)  sleep "$RUN_SLEEP_SEC" ;;
    SU)     break ;;
    FA|TE|IN|OH|OI|'') sleep "$RETRY_SLEEP_SEC" ;;
    *)      sleep "$RETRY_SLEEP_SEC" ;;
  esac
done

while : ; do
  mins="$(mins_until_next_start)"
  if [ "$mins" -le "$CUTOFF_BUFFER_MIN" ]; then exit 0; fi
  BDATES="$(cat "$STATE_FILE")"
  [ -n "$BDATES" ] || break
  first="${BDATES%% *}"
  rest="${BDATES#* }"
  if [ "$rest" = "$BDATES" ]; then rest=""; fi
  force_start_now
  wait_until_su
  echo "$rest" > "$STATE_FILE"
  [ -n "$first" ] && printf "%s\n" "$first" >> "$HISTORY_FILE"
done

exit 0