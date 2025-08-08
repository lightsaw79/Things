#!/bin/ksh
# -----------------------------------------------------------------------------
# autosys_controller.ksh — ORG STYLE, EXTREME DEBUG (ksh88/pdksh safe)
# Arg order (per your JIL):
#   $1 BOX_NAME
#   $2 RESET_JOB_NAME
#   $3 GET_STATUS_JOB_NAME
#   $4 APP_ID
#   $5 BUS_DATE_PARAM
#
# Flow:
#   1) autosys_get_status.ksh <BOX> <APP_ID> <BUS_DATE> [...]
#   2) autosys_reset.ksh     <BOX> <APP_ID> <BUS_DATE> [...] -> INACTIVE
#   3) ${AUTOSYS_RESET_DIR}/sendevent -P 1 -E STARTJOB -J <BOX>
#
# Exit codes:
#   0 OK | 11 get_status failed | 12 reset failed | 13 start failed
#   90 ENV missing | 91/92/93 binary missing | 97 args missing | 98 checker flagged
# -----------------------------------------------------------------------------

# Optional xtrace (set in job profile: CTRL_TRACE=1)
[ "${CTRL_TRACE:-0}" = "1" ] && set -x
set +e   # never inherit set -e

# --- Environment (ORG STANDARD) ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "ERROR: ENV_FILE not found: $ENV_FILE" >&2
  exit 90
fi
. "$ENV_FILE"

# --- Arguments (NEW ORDER) ---
BOX_NAME=$1
RESET_JOB_NAME=$2
GET_STATUS_JOB_NAME=$3
APP_ID=$4
BUS_DATE_PARAM=$5
CATEGORY_ID=1

# --- Optional knobs (override in job profile if needed) ---
[ -z "$CTRL_ENABLE_CHECKER" ] && CTRL_ENABLE_CHECKER=1
[ -z "$CTRL_GREPPAT" ] && CTRL_GREPPAT='(ORA-[0-9]+|SQLSTATE|^ERROR| ERROR |FATAL|EXCEPTION|Abend|Segmentation fault|command not found|JOBFAILURE)'
[ -z "$CTRL_WHITELIST" ] && CTRL_WHITELIST='No errors found|0 errors|RETURN CODE:0|NOT an error|Checking Errors in File|Attaching File|Mail Sent'
[ -z "$CTRL_HEAD" ] && CTRL_HEAD=60
[ -z "$CTRL_TAIL" ] && CTRL_TAIL=60

# --- Validate args early ---
if [ -z "$BOX_NAME" ] || [ -z "$RESET_JOB_NAME" ] || [ -z "$GET_STATUS_JOB_NAME" ] || [ -z "$APP_ID" ] || [ -z "$BUS_DATE_PARAM" ]; then
  echo "USAGE: $0 <BOX_NAME> <RESET_JOB_NAME> <GET_STATUS_JOB_NAME> <APP_ID> <BUS_DATE_PARAM>" >&2
  exit 97
fi

# --- Runtime / names ---
SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
HOSTNAME_FQDN=`hostname 2>/dev/null || uname -n`
RUN_USER=`id -un 2>/dev/null || whoami`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

# --- Logs ---
CTRL_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
CTRL_ERR="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err.log"
START_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_start_${TIMESTAMP}.log"

# --- Paths from env ---
GET_STATUS_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_get_status.ksh
RESET_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_reset.ksh
SENDEVENT_BIN=${AUTOSYS_RESET_DIR}/sendevent
AUTOREP_WRAPPER=${AUTOSYS_RESET_DIR}/autorep.sh
RAW_AUTOREP=/opt/CA/WorkloadAutomationAE/autosys/bin/autorep

# --- Helpers ---
now() { date '+%Y-%m-%d %H:%M:%S'; }
log() { echo "`now` | $SCRIPT_NAME | $*"; }
logf(){ log "$*" | tee -a "$CTRL_LOG"; }
fail(){ log "ERROR: $*" | tee -a "$CTRL_ERR" 1>&2; }

run_cmd() {
  _LABEL="$1"; shift
  _START=`date +%s`
  logf "[RUN] ${_LABEL}: $*"
  "$@" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
  _RC=$?
  _END=`date +%s`
  _DUR=`expr $_END - $_START`
  if [ $_RC -eq 0 ]; then
    logf "[OK ] ${_LABEL} | RC=$_RC | ${_DUR}s"
  else
    fail "[BAD] ${_LABEL} | RC=$_RC | ${_DUR}s"
  fi
  return $_RC
}

safe_fx_event() { fx_sam8_odm_event_log "$@"; _RC=$?; [ $_RC -ne 0 ] && logf "WARN: fx_sam8_odm_event_log RC=$_RC"; }
safe_chk_start(){ CHKDOMBATCHSCHEDULESTART "$@"; _RC=$?; [ $_RC -ne 0 ] && logf "WARN: CHKDOMBATCHSCHEDULESTART RC=$_RC"; }
safe_chk_end()  { CHKDOMBATCHSCHEDULEEND   "$@"; _RC=$?; [ $_RC -ne 0 ] && logf "WARN: CHKDOMBATCHSCHEDULEEND RC=$_RC"; }

get_box_status() {
  if [ -x "$AUTOREP_WRAPPER" ]; then
    "$AUTOREP_WRAPPER" -q "$BOX_NAME" 2>>"$CTRL_ERR" | awk 'NR==1{print $2}'
  else
    "$RAW_AUTOREP" -j "$BOX_NAME" -q 2>>"$CTRL_ERR" | awk 'NR==1{print $2}'
  fi
}

dump_file_info() {
  _F="$1"; [ -n "$_F" ] || return 0
  {
    echo "---- FILE INFO: $_F ----"
    ls -l $_F 2>&1
    file $_F 2>&1
    head -n 5 $_F 2>/dev/null || true
    echo "---- END FILE INFO ----"
  } >>"$CTRL_LOG"
}

# --- Quote/escape helper for ODM message strings (bullet-proof) ---
_esc() {
  _TMP=`echo "$1" | sed "s/'/''/g"`       # single quotes doubled
  _TMP=`echo "$_TMP" | sed "s/&/\\\&/g"`  # ampersand escaped for SQL*Plus
  echo "$_TMP"
}
MSG_START="'Execution Started $(_esc "$SCRIPT_NAME") for $(_esc "$BOX_NAME")'"
MSG_END="'Execution End $(_esc "$SCRIPT_NAME") for $(_esc "$BOX_NAME")'"

# --- DEBUG HEADER ---
{
  echo "===== DEBUG HEADER ====="
  echo "SCRIPT_NAME=$SCRIPT_NAME"
  echo "HOST=$HOSTNAME_FQDN  USER=$RUN_USER  UMASK=`umask`"
  echo "DATE=`now`"
  echo "BOX_NAME=$BOX_NAME"
  echo "RESET_JOB_NAME=$RESET_JOB_NAME"
  echo "GET_STATUS_JOB_NAME=$GET_STATUS_JOB_NAME"
  echo "APP_ID=$APP_ID  BUS_DATE=$BUS_DATE_PARAM"
  echo "AUTOSYS_COMMON_BIN_DIR=$AUTOSYS_COMMON_BIN_DIR"
  echo "AUTOSYS_RESET_DIR=$AUTOSYS_RESET_DIR"
  echo "GET_STATUS_SH=$GET_STATUS_SH"
  echo "RESET_SH=$RESET_SH"
  echo "SENDEVENT_BIN=$SENDEVENT_BIN"
  echo "AUTOREP_WRAPPER=$AUTOREP_WRAPPER"
  echo "RAW_AUTOREP=$RAW_AUTOREP"
  echo "PATH=$PATH"
  echo "CTRL_ENABLE_CHECKER=$CTRL_ENABLE_CHECKER"
  echo "CTRL_GREPPAT=$CTRL_GREPPAT"
  echo "CTRL_WHITELIST=$CTRL_WHITELIST"
  echo "========================"
} | tee -a "$CTRL_LOG"

# --- Existence checks ---
[ -x "$GET_STATUS_SH" ] || { fail "missing $GET_STATUS_SH"; dump_file_info "$GET_STATUS_SH"; exit 91; }
[ -x "$RESET_SH"     ] || { fail "missing $RESET_SH";     dump_file_info "$RESET_SH";     exit 92; }
[ -x "$SENDEVENT_BIN" ] || { fail "missing $SENDEVENT_BIN";                                 exit 93; }

dump_file_info "$GET_STATUS_SH"
dump_file_info "$RESET_SH"
dump_file_info "$SENDEVENT_BIN"

# --- STDERR tee using named pipe (ksh88-safe) ---
CTRL_ERR_PIPE="/tmp/${SCRIPT_NAME}_$$.errpipe"
mkfifo "$CTRL_ERR_PIPE"
tee -a "$CTRL_ERR" <"$CTRL_ERR_PIPE" >&2 &
TEE_PID=$!
exec 2>"$CTRL_ERR_PIPE"

# --- Mark start time for ExecSql log scoping ---
CTRL_MARK="/tmp/${SCRIPT_NAME}_$$.mark"
: > "$CTRL_MARK"

# --- Audit start (non-fatal) ---
safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 $MSG_START
safe_chk_start "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

# --- Status BEFORE ---
CUR_ST=`get_box_status`
logf "CURRENT BOX STATUS [$BOX_NAME] = ${CUR_ST:-UNKNOWN}"

# --- 1) GET_STATUS ---
run_cmd "get_status" "$GET_STATUS_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
RC=$?; if [ $RC -ne 0 ]; then
  fail "autosys_get_status.ksh failed (RC=$RC)"
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 $MSG_END
  exec 2>&1; kill $TEE_PID 2>/dev/null; rm -f "$CTRL_ERR_PIPE" "$CTRL_MARK"; exit 11
fi

# --- 2) RESET ---
run_cmd "reset" "$RESET_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
RC=$?; if [ $RC -ne 0 ]; then
  fail "autosys_reset.ksh failed (RC=$RC)"
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 $MSG_END
  exec 2>&1; kill $TEE_PID 2>/dev/null; rm -f "$CTRL_ERR_PIPE" "$CTRL_MARK"; exit 12
fi

POST_RESET_ST=`get_box_status`
logf "POST-RESET BOX STATUS [$BOX_NAME] = ${POST_RESET_ST:-UNKNOWN}"

# --- 3) START BOX NOW ---
run_cmd "start_box" "$SENDEVENT_BIN" -P 1 -E STARTJOB -J "${BOX_NAME}"
RC=$?; if [ $RC -ne 0 ]; then
  fail "STARTJOB ${BOX_NAME} failed (RC=$RC)"
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 $MSG_END
  exec 2>&1; kill $TEE_PID 2>/dev/null; rm -f "$CTRL_ERR_PIPE" "$CTRL_MARK"; exit 13
fi

AFTER_START_ST=`get_box_status`
logf "AFTER-START BOX STATUS [$BOX_NAME] = ${AFTER_START_ST:-UNKNOWN}"

# --- ExecSql checker: only scan logs created during THIS run ---
_CE_RC=0
if [ "$CTRL_ENABLE_CHECKER" = "1" ]; then
  LATEST_EXECLOG=`find /apps/samd/actimize/package_utilities/common/Log \
                    -type f -name 'ExecSql.*.log' -newer "$CTRL_MARK" \
                    -printf '%T@ %p\n' 2>/dev/null | sort -nr | head -1 | awk '{ $1=""; sub(/^ /,""); print }'`

  if [ -n "$LATEST_EXECLOG" ] && [ -f "$LATEST_EXECLOG" ]; then
    logf "CHECK_ERROR: Scanning ExecSql from this run: $LATEST_EXECLOG"
    logf "CHECK_ERROR: Pattern: $CTRL_GREPPAT"
    logf "CHECK_ERROR: Whitelist: $CTRL_WHITELIST"

    echo "----- ExecSql HEAD (first $CTRL_HEAD) -----" | tee -a "$CTRL_LOG"
    head -n $CTRL_HEAD "$LATEST_EXECLOG" 2>&1 | tee -a "$CTRL_LOG" >>"$CTRL_ERR"
    echo "----- ExecSql TAIL (last $CTRL_TAIL) -----"  | tee -a "$CTRL_LOG"
    tail -n $CTRL_TAIL "$LATEST_EXECLOG" 2>&1 | tee -a "$CTRL_LOG" >>"$CTRL_ERR"
    echo "----- ExecSql MATCHES (grep -nE) -----"      | tee -a "$CTRL_LOG"

    LC_ALL=C grep -nE "$CTRL_GREPPAT" "$LATEST_EXECLOG" > /tmp/_ctrl_matches.$$ 2>>"$CTRL_ERR"
    if [ -s /tmp/_ctrl_matches.$$ ]; then
      cat /tmp/_ctrl_matches.$$ | tee -a "$CTRL_LOG"
      if [ -n "$CTRL_WHITELIST" ]; then
        LC_ALL=C egrep -vi "$CTRL_WHITELIST" /tmp/_ctrl_matches.$$ > /tmp/_ctrl_real.$$
        if [ -s /tmp/_ctrl_real.$$ ]; then
          fail "CHECK_ERROR: Non-whitelisted matches found (see above)"
          _CE_RC=98
        else
          logf "CHECK_ERROR: Only whitelisted matches; treating as OK."
          _CE_RC=0
        fi
        rm -f /tmp/_ctrl_real.$$
      else
        fail "CHECK_ERROR: Matches found (no whitelist applied)"
        _CE_RC=98
      fi
    else
      echo "NO MATCHES" | tee -a "$CTRL_LOG"
      _CE_RC=0
    fi
    echo "----- End ExecSql Debug -----" | tee -a "$CTRL_LOG"
    rm -f /tmp/_ctrl_matches.$$
  else
    logf "CHECK_ERROR: No ExecSql log created during this run; skipping scan."
    _CE_RC=0
  fi
else
  logf "CHECK_ERROR: Disabled by CTRL_ENABLE_CHECKER=0"
  _CE_RC=0
fi

# --- Audit end + cleanup ---
safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 $MSG_END

# close stderr tee cleanly
exec 2>&1
kill $TEE_PID 2>/dev/null
rm -f "$CTRL_ERR_PIPE" "$CTRL_MARK"

[ ${_CE_RC:-0} -eq 0 ] || exit $_CE_RC
exit 0