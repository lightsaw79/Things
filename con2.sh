#!/bin/ksh
# -----------------------------------------------------------------------------
# autosys_hourly_controller.ksh
# Controller job to:
#   1) Call autosys_get_status.ksh for <BOX_NAME>
#   2) Call autosys_reset.ksh for <BOX_NAME> (sets box INACTIVE)
#   3) Force start the box immediately
# -----------------------------------------------------------------------------

# --- Environment file check ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "ERROR: ENV_FILE not found: $ENV_FILE" >&2
  exit 90
fi
. "$ENV_FILE"

# --- Arguments ---
BOX_NAME=$1
APP_ID=$2
BUS_DATE_PARAM=$3
CATEGORY_ID=1

SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

# --- Logs ---
CTRL_LOG="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
CTRL_ERR="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err.log"

# --- Paths ---
GET_STATUS_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_get_status.ksh
RESET_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_reset.ksh
SENDEVENT_BIN=${AUTOSYS_RESET_DIR}/sendevent

# --- Debug & binary existence patch ---
echo "DEBUG ENV: AUTOSYS_COMMON_BIN_DIR=$AUTOSYS_COMMON_BIN_DIR" | tee -a "$CTRL_LOG"
echo "DEBUG ENV: AUTOSYS_RESET_DIR=$AUTOSYS_RESET_DIR"           | tee -a "$CTRL_LOG"
echo "DEBUG PATHS: GET_STATUS_SH=$GET_STATUS_SH"                 | tee -a "$CTRL_LOG"
echo "DEBUG PATHS: RESET_SH=$RESET_SH"                           | tee -a "$CTRL_LOG"

[ -x "$GET_STATUS_SH" ] || { echo "ERROR: missing $GET_STATUS_SH" | tee -a "$CTRL_ERR"; exit 91; }
[ -x "$RESET_SH"     ] || { echo "ERROR: missing $RESET_SH"       | tee -a "$CTRL_ERR"; exit 92; }
[ -x "$SENDEVENT_BIN" ] || { echo "ERROR: missing $SENDEVENT_BIN" | tee -a "$CTRL_ERR"; exit 93; }

# --- Audit start ---
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 "Execution Started $SCRIPT_NAME for $BOX_NAME"
CHKDOMBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

# --- 1) Get status ---
echo "`date '+%Y-%m-%d %H:%M:%S'` | Calling autosys_get_status.ksh ..." | tee -a "$CTRL_LOG"
$GET_STATUS_SH "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | autosys_get_status.ksh failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 11
fi

# --- 2) Reset ---
echo "`date '+%Y-%m-%d %H:%M:%S'` | Calling autosys_reset.ksh ..." | tee -a "$CTRL_LOG"
$RESET_SH "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM" >>"$CTRL_LOG" 2>>"$CTRL_ERR"
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | autosys_reset.ksh failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 12
fi

# --- 3) Start the box ---
echo "`date '+%Y-%m-%d %H:%M:%S'` | Submitting START for [$BOX_NAME]" | tee -a "$CTRL_LOG"
$SENDEVENT_BIN -P 1 -E STARTJOB -J ${BOX_NAME} > ${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_start_${TIMESTAMP}.log 2>&1
RC=$?
if [ $RC -ne 0 ]; then
  LogError $APP_ID $RC "$SCRIPT_NAME | STARTJOB ${BOX_NAME} failed (RC=$RC)"
  CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
  exit 13
fi
echo "`date '+%Y-%m-%d %H:%M:%S'` | STARTJOB submitted for [$BOX_NAME]" | tee -a "$CTRL_LOG"

# --- Audit end ---
CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End $SCRIPT_NAME for $BOX_NAME"

exit 0