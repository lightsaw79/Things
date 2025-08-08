#!/bin/ksh
# -----------------------------------------------------------------------------
# autosys_controller.ksh  (ORG STYLE, EXTREME DEBUG)
# Flow (hourly, controller owns cadence):
#   1) autosys_get_status.ksh <BOX> <APP_ID> <BUS_DATE> [GET_EXTRA]
#   2) autosys_reset.ksh     <BOX> <APP_ID> <BUS_DATE> [RESET_EXTRA]  -> INACTIVE
#   3) ${AUTOSYS_RESET_DIR}/sendevent -P 1 -E STARTJOB -J <BOX>       -> start now
# Exits: 90 env missing, 91/92/93 missing binaries, 11/12/13 step fails, 0 OK
# -----------------------------------------------------------------------------

# Be strict but don’t die from errexit in sourced env
set +e

# --- Environment (org standard) ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "ERROR: ENV_FILE not found: $ENV_FILE" >&2
  exit 90
fi
. "$ENV_FILE"

# --- Args ---
BOX_NAME=$1
APP_ID=$2
BUS_DATE_PARAM=$3
GET_EXTRA="$4"     # optional: extra flags for get_status
RESET_EXTRA="$5"   # optional: extra flags for reset
CATEGORY_ID=1

# --- Runtime / log names ---
SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
HOSTNAME_FQDN=`hostname 2>/dev/null || uname -n`
RUN_USER=`id -un 2>/dev/null || whoami`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

CTRL_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
CTRL_ERR="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err.log"
START_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_start_${TIMESTAMP}.log"

# --- Paths from env (org conventions) ---
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

safe_fx_event() {
  fx_sam8_odm_event_log "$@" ; _RC=$?
  [ $_RC -ne 0 ] && logf "WARN: fx_sam8_odm_event_log RC=$_RC (continuing)"
  return 0
}
safe_chk_start() {
  CHKDOMBATCHSCHEDULESTART "$@" ; _RC=$?
  [ $_RC -ne 0 ] && logf "WARN: CHKDOMBATCHSCHEDULESTART RC=$_RC (continuing)"
  return 0
}
safe_chk_end() {
  CHKDOMBATCHSCHEDULEEND "$@" ; _RC=$?
  [ $_RC -ne 0 ] && logf "WARN: CHKDOMBATCHSCHEDULEEND RC=$_RC (continuing)"
  return 0
}

get_box_status() {
  if [ -x "$AUTOREP_WRAPPER" ]; then
    "$AUTOREP_WRAPPER" -q "$BOX_NAME" 2>>"$CTRL_ERR" | awk 'NR==1{print $2}'
  else
    "$RAW_AUTOREP" -j "$BOX_NAME" -q 2>>"$CTRL_ERR" | awk 'NR==1{print $2}'
  fi
}

dump_file_info() {
  _F="$1"
  [ -n "$_F" ] || return 0
  {
    echo "---- FILE INFO: $_F ----"
    ls -l $_F 2>&1
    file $_F 2>&1
    head -n 5 $_F 2>/dev/null || true
    echo "---- END FILE INFO ----"
  } >>"$CTRL_LOG"
}

# --- SUPER DEBUG HEADER ---
{
  echo "===== DEBUG HEADER ====="
  echo "SCRIPT_NAME=$SCRIPT_NAME"
  echo "HOST=$HOSTNAME_FQDN  USER=$RUN_USER  UMASK=`umask`"
  echo "DATE=`now`"
  echo "BOX_NAME=$BOX_NAME  APP_ID=$APP_ID  BUS_DATE=$BUS_DATE_PARAM"
  echo "GET_EXTRA=$GET_EXTRA"
  echo "RESET_EXTRA=$RESET_EXTRA"
  echo "AUTOSYS_COMMON_BIN_DIR=$AUTOSYS_COMMON_BIN_DIR"
  echo "AUTOSYS_RESET_DIR=$AUTOSYS_RESET_DIR"
  echo "GET_STATUS_SH=$GET_STATUS_SH"
  echo "RESET_SH=$RESET_SH"
  echo "SENDEVENT_BIN=$SENDEVENT_BIN"
  echo "AUTOREP_WRAPPER=$AUTOREP_WRAPPER"
  echo "RAW_AUTOREP=$RAW_AUTOREP"
  echo "PATH=$PATH"
  echo "========================"
} | tee -a "$CTRL_LOG"

# --- Existence checks (fail fast with details) ---
[ -x "$GET_STATUS_SH" ] || { fail "missing $GET_STATUS_SH"; dump_file_info "$GET_STATUS_SH"; exit 91; }
[ -x "$RESET_SH"     ] || { fail "missing $RESET_SH";     dump_file_info "$RESET_SH";     exit 92; }
[ -x "$SENDEVENT_BIN" ] || { fail "missing $SENDEVENT_BIN";                                 exit 93; }

dump_file_info "$GET_STATUS_SH"
dump_file_info "$RESET_SH"
dump_file_info "$SENDEVENT_BIN"

# --- Audit start (non-fatal wrappers) ---
safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 "Execution Started $SCRIPT_NAME for $BOX_NAME"
safe_chk_start "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

# --- Status before actions ---
CUR_ST=`get_box_status`
logf "CURRENT BOX STATUS [$BOX_NAME] = ${CUR_ST:-UNKNOWN}"

# --- 1) get_status ---
if [ -n "$GET_EXTRA" ]; then
  run_cmd "get_status" "$GET_STATUS_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" $GET_EXTRA
else
  run_cmd "get_status" "$GET_STATUS_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
fi
RC=$?
if [ $RC -ne 0 ]; then
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End (get_status failed RC=$RC) $SCRIPT_NAME for $BOX_NAME"
  exit 11
fi

# --- 2) reset ---
if [ -n "$RESET_EXTRA" ]; then
  run_cmd "reset" "$RESET_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" $RESET_EXTRA
else
  run_cmd "reset" "$RESET_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
fi
RC=$?
if [ $RC -ne 0 ]; then
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End (reset failed RC=$RC) $SCRIPT_NAME for $BOX_NAME"
  exit 12
fi

POST_RESET_ST=`get_box_status`
logf "POST-RESET BOX STATUS [$BOX_NAME] = ${POST_RESET_ST:-UNKNOWN}"

# --- 3) Start box now (event-driven) ---
run_cmd "start_box" "$SENDEVENT_BIN" -P 1 -E STARTJOB -J "${BOX_NAME}"
RC=$?
if [ $RC -ne 0 ]; then
  safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
  safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End (start failed RC=$RC) $SCRIPT_NAME for $BOX_NAME"
  exit 13
fi

AFTER_START_ST=`get_box_status`
logf "AFTER-START BOX STATUS [$BOX_NAME] = ${AFTER_START_ST:-UNKNOWN}"

# --- CHECK_ERROR style deep dive: show WHY RC=1 happens post-run ---
# Find latest ExecSql log your org scanner looks at
LATEST_EXECLOG=`ls -1t /apps/samd/actimize/package_utilities/common/Log/ExecSql.*.log 2>/dev/null | head -1`
if [ -n "$LATEST_EXECLOG" ]; then
  logf "CHECK_ERROR: Latest ExecSql log is: $LATEST_EXECLOG"
  dump_file_info "$LATEST_EXECLOG"

  # Patterns your org typically trips on (tweak if needed)
  CE_PAT='(ORA-[0-9]+|SQLSTATE|^ERROR| ERROR |FATAL|EXCEPTION|Traceback|Abend|Segmentation fault|command not found|RETURN CODE:[1-9]|JOBFAILURE)'
  logf "CHECK_ERROR: Using grep -nE pattern: $CE_PAT"

  # Show first/last lines and *all* matching lines with line numbers
  {
    echo "----- ExecSql HEAD (first 80 lines) -----"
    head -n 80 "$LATEST_EXECLOG" 2>&1
    echo "----- ExecSql TAIL (last 80 lines) -----"
    tail -n 80 "$LATEST_EXECLOG" 2>&1
    echo "----- ExecSql MATCHES (grep -nE) -----"
    LC_ALL=C grep -nE "$CE_PAT" "$LATEST_EXECLOG" || echo "NO MATCHES"
    echo "----- End ExecSql Debug -----"
  } >>"$CTRL_LOG" 2>>"$CTRL_ERR"
else
  logf "CHECK_ERROR: No ExecSql.*.log found to scan."
fi

# --- Audit end ---
safe_chk_end "$APP_ID" "$BUS_DATE_PARAM"
safe_fx_event $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End $SCRIPT_NAME for $BOX_NAME"

exit 0