#!/bin/ksh
# controller_inline_from_table.ksh
# Usage: controller_inline_from_table.ksh <BOX_NAME>
# Optional: BATCH_TABLE_FILE=/path/to/your_table.txt controller_inline_from_table.ksh <BOX_NAME>

box_name="$1"

LOGDIR="${MODEL_BATCH_LOGFILES_DIR:-/tmp}"
STATE_FILE="${LOGDIR}/${box_name}_bdates.lst"

# ---- locate the table file on your UNIX server (super lean) ----
TFILE="${BATCH_TABLE_FILE:-}"
if [ -z "$TFILE" ]; then
  for d in /apps/samd/actimize/data /apps/samd/actimize/output /apps/samd/actimize/logs
  do
    TFILE=$(ls -1t $d/*[Bb][Aa][Tt][Cc][Hh]*.{txt,csv,log} 2>/dev/null | head -1)
    [ -n "$TFILE" ] && break
  done
fi

# If still not found or empty, do nothing this run
[ -s "$TFILE" ] || exit 0

# ---- build (or load) pending BD list from *your* table ----
# Keep ETL rows with COMPLETED that have NO matching COMPLETED SAM with same APP_ID.
if [ -s "$STATE_FILE" ]; then
  BDATES="$(cat "$STATE_FILE")"
else
  BDATES="$(awk '
    BEGIN{ IGNORECASE=0 }
    function is_date(x){ return x ~ /^[0-3]?[0-9]-[A-Z]{3}-[0-9]{2,4}$/ }

    # Record COMPLETED SAM app_ids
    /_SAM_BATCH/ && $0 ~ /COMPLETED/ { sam[$2]=1; next }

    # Record COMPLETED ETL app_ids + their business date (last date-like token)
    /_ETL_BATCH/ && $0 ~ /COMPLETED/ {
      d=""
      for (i=NF; i>=1; i--) if (is_date($i)) { d=$i; break }
      if (d!="") { bd[$2]=d; order[++n]=$2 }
    }

    END {
      first=1
      for(i=1;i<=n;i++){
        id=order[i]
        if(!(id in sam)){
          if(!first) printf " "
          printf "%s", bd[id]      # keep duplicates, preserve ETL order
          first=0
        }
      }
    }' "$TFILE")"
  echo "$BDATES" > "$STATE_FILE"
fi

# ---- status (prefer your org wrapper) ----
get_status() {
  if [ -x "${AUTOSYS_RESET_DIR}/autorep.sh" ]; then
    "${AUTOSYS_RESET_DIR}/autorep.sh" -q "$box_name" 2>/dev/null | awk 'NR==2{print $3}'
  else
    autorep -J "$box_name" -q 2>/dev/null | awk 'NR==2{print $3}'
  fi
}

# ---- inline force-start (your exact style) ----
force_start_now() {
  "${AUTOSYS_RESET_DIR}/sendevent" -p 1 -E CHANGE_STATUS -s INACTIVE -J "${box_name}" \
    > "${LOGDIR}/${box_name}_inactive.log" 2>&1
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "${box_name}" \
    > "${LOGDIR}/${box_name}.log" 2>&1
}

# ---- main ----
STATUS="$(get_status)"
case "$STATUS" in
  RUNNING|ACTIVATED)
    exit 0
    ;;
  SUCCESS)
    if [ -n "$BDATES" ]; then
      REST="$(echo "$BDATES" | sed 's/^[^ ]*[ ]*//')"   # pop first date
      force_start_now
      echo "$REST" > "$STATE_FILE"
    else
      rm -f "$STATE_FILE" >/dev/null 2>&1
    fi
    exit 0
    ;;
  FAILURE|TERMINATED|INACTIVE|ON_HOLD|ON_ICE|'*'|'')
    force_start_now
    exit 0
    ;;
  *)
    force_start_now
    exit 0
    ;;
esac