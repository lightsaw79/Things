#!/bin/ksh
# Controller: every run -> get_status -> reset -> START box now (org style)

# --- Environment (org style) ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "Environment file [$ENV_FILE] not found. Exiting ..."
  exit 1
fi
. $ENV_FILE

# Args: BOX_NAME APP_ID BUS_DATE_PARAM
BOX_NAME=$1
APP_ID=$2
BUS_DATE_PARAM=$3
CATEGORY_ID=1

SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

# Logs (org locations)
CTRL_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
CTRL_ERR="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err.log"

log()     { echo "`date '+%Y-%m-%d %H:%M:%S'` | $SCRIPT_NAME | $*"; }
logf()    { log "$*" | tee -a "$CTRL_LOG"; }
logerrf() { log "ERROR: $*" | tee -a "$CTRL_ERR" 1>&2; }

# Your org utilities
GET_STATUS_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_get_status.ksh
RESET_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_reset.ksh

# sendevent exactly like reset.ksh (same var + -P 1)
SENDEVENT_BIN=${AUTOSYS_RESET_DIR}/sendevent

# --- Audit start + window guard ---
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 "Execution Started $SCRIPT_NAME for $BOX_NAME"
CHKDOMBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

logf "DEBUG: BOX=$BOX_NAME APP_ID=$APP_ID BD=$BUS_DATE_PARAM"

# 1) get_status (org utility)
logf "Calling autosys_get_status.ksh ..."
$GET_STATUS_SH "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | autosys_get_status.ksh failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 11
fi

# 2) reset (org utility; sets box INACTIVE internally)
logf "Calling autosys_reset.ksh ..."
$RESET_SH "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | autosys_reset.ksh failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 12
fi

# 3) Start the box NOW (no calendars) — org-style sendevent line
logf "Submitting START for [$BOX_NAME]"
$SENDEVENT_BIN -P 1 -E STARTJOB -J ${BOX_NAME} > ${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_start_${TIMESTAMP}.log 2>&1
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | STARTJOB ${BOX_NAME} failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 13
fi
logf "STARTJOB submitted for [$BOX_NAME]"

# --- Close guard + audit end ---
CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End $SCRIPT_NAME for $BOX_NAME"
exit 0