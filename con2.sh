#!/bin/ksh

# --- Environment ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "Environment file [$ENV_FILE] not found. Exiting ..."
  exit 1
fi
. $ENV_FILE

# --- Args from JIL (NO hard-coding) ---
BOX_NAME=$1
GET_STATUS_JOB=$2
RESET_JOB=$3
APP_ID=$4
BUS_DATE_PARAM=$5
CATEGORY_ID=1

# --- Runtime / names ---
SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

# --- Org wrappers / tools (fallback to raw if wrappers missing) ---
AUTOREP_WRAPPER=${AUTOSYS_RESET_DIR}/autorep.sh
RAW_AUTOREP=/opt/CA/WorkloadAutomationAE/autosys/bin/autorep

SENDEVENT_BIN=${AUTOSYS_RESET_DIR}/sendevent
[ -x "$SENDEVENT_BIN" ] || SENDEVENT_BIN=/opt/CA/WorkloadAutomationAE/autosys/bin/sendevent

# --- Local logs (same pattern as other utilities) ---
CTRL_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
CTRL_ERR="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err.log"

# --- Helpers (echo-only) ---
log()     { echo "`date '+%Y-%m-%d %H:%M:%S'` | $SCRIPT_NAME | $*"; }
logf()    { log "$*"       | tee -a "$CTRL_LOG" ; }
logerrf() { log "ERROR: $*" | tee -a "$CTRL_ERR" 1>&2; }

get_status() {
  if [ -x "$AUTOREP_WRAPPER" ]; then
    $AUTOREP_WRAPPER -q "$1" 2>/dev/null | awk 'NR==1{print $2}'
  else
    $RAW_AUTOREP -j "$1" -q 2>/dev/null | awk 'NR==1{print $2}'
  fi
}

wait_for_job_end() {
  _JOB="$1"; _TIMEOUT="$2"; _i=0
  while : ; do
    _st=`get_status "$_JOB"`
    case "$_st" in
      SU|FA|TE) logf "$_JOB completed with status=$_st"; return 0 ;;
    esac
    _i=`expr $_i + 1`
    [ $_i -ge $_TIMEOUT ] && { logerrf "Timeout waiting for $_JOB to end (last=$_st)"; return 1; }
    sleep 5
  done
}

wait_for_box_inactive() {
  _BOX="$1"; _TIMEOUT="$2"; _i=0
  while : ; do
    _st=`get_status "$_BOX"`
    [ "$_st" = "INACTIVE" ] && { logf "Box $_BOX is INACTIVE"; return 0; }
    _i=`expr $_i + 1`
    [ $_i -ge $_TIMEOUT ] && { logerrf "Timeout waiting for $_BOX to become INACTIVE (last=$_st)"; return 1; }
    sleep 5
  done
}

send_force_start() {
  _OBJ="$1"
  $SENDEVENT_BIN -E FORCE_STARTJOB -J "$_OBJ" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
  _rc=$?
  if [ $_rc -ne 0 ]; then
    LogError $APP_ID $_rc "$SCRIPT_NAME | FORCE_START $_OBJ failed (RC=$_rc)"
    logerrf "FORCE_START $_OBJ failed (RC=$_rc)"
    return $_rc
  else
    Log "$SCRIPT_NAME | FORCE_START $_OBJ OK (RC=$_rc)"
    logf "FORCE_START $_OBJ OK"
    return 0
  fi
}

# --- Start audit + SLA guard (matches your other scripts) ---
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 "Execution Started $SCRIPT_NAME for $BOX_NAME"
# Keep trap commented while stabilizing; re-enable when steady
# trap "LogError $APP_ID $? '$SCRIPT_NAME trapped signal'" ERR 2 3 15
CHKDOMBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

# --- DEBUG banner ---
Log "$SCRIPT_NAME | DEBUG: BOX=$BOX_NAME GET=$GET_STATUS_JOB RESET=$RESET_JOB APP_ID=$APP_ID SENDEVENT_BIN=$SENDEVENT_BIN AUTOREP_WRAPPER=$AUTOREP_WRAPPER"
logf "DEBUG: args=[$BOX_NAME] [$GET_STATUS_JOB] [$RESET_JOB] APP_ID=[$APP_ID] BD=[$BUS_DATE_PARAM]"

# --- Controller logic ---
CUR_ST=`get_status "$BOX_NAME"`
Log "$SCRIPT_NAME | Current status of [$BOX_NAME] = $CUR_ST"
logf "Current status of [$BOX_NAME] = $CUR_ST"

if [ "$CUR_ST" != "SU" ]; then
  # 1) get_status job
  Log "$SCRIPT_NAME | Launching $GET_STATUS_JOB"
  send_force_start "$GET_STATUS_JOB" || { CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 11; }
  wait_for_job_end "$GET_STATUS_JOB" 120 || { LogError $APP_ID 11 "$SCRIPT_NAME | $GET_STATUS_JOB timeout"; CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 11; }

  # 2) reset job
  Log "$SCRIPT_NAME | Launching $RESET_JOB"
  send_force_start "$RESET_JOB" || { CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 12; }
  wait_for_job_end "$RESET_JOB" 240     || { LogError $APP_ID 12 "$SCRIPT_NAME | $RESET_JOB timeout"; CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 12; }
  wait_for_box_inactive "$BOX_NAME" 240 || { LogError $APP_ID 13 "$SCRIPT_NAME | Box INACTIVE wait timeout"; CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 13; }
else
  Log "$SCRIPT_NAME | Box already SU; no heal needed"
  logf "Box is SU; skipping reset"
fi

# 3) Kick next cycle
Log "$SCRIPT_NAME | Kicking next cycle for $BOX_NAME"
send_force_start "$BOX_NAME" || { CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"; exit 14; }

# --- Close SLA window + audit ---
CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End $SCRIPT_NAME for $BOX_NAME"
exit 0