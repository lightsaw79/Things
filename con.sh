#!/bin/ksh
# -----------------------------------------------------------------------------
# LEAN autosys_controller.ksh (ksh88-safe, minimal + RC breadcrumbs)
# Args (per JIL):
#   $1 BOX_NAME
#   $2 RESET_JOB_NAME               # not used for calls; kept for log context
#   $3 GET_STATUS_JOB_NAME          # not used for calls; kept for log context
#   $4 APP_ID
#   $5 BUS_DATE_PARAM
#
# Flow:
#   1) ${AUTOSYS_COMMON_BIN_DIR}/autosys_get_status.ksh <BOX> <APP_ID> <BUS_DATE>
#   2) ${AUTOSYS_COMMON_BIN_DIR}/autosys_reset.ksh     <BOX> <APP_ID> <BUS_DATE>
#   3) ${AUTOSYS_RESET_DIR}/sendevent -P 1 -E STARTJOB -J <BOX>
#
# Exit codes: 0 OK | 11 get_status failed | 12 reset failed | 13 start failed
#             90 env missing | 91/92/93 missing binaries | 97 bad args
# -----------------------------------------------------------------------------

set +e   # handle RCs ourselves

# --- Org environment ---
export ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  echo "ERROR: ENV file not found: $ENV_FILE" >&2
  exit 90
fi
. "$ENV_FILE"

# --- Args (your order) ---
BOX_NAME=$1
RESET_JOB_NAME=$2
GET_STATUS_JOB_NAME=$3
APP_ID=$4
BUS_DATE_PARAM=$5

if [ -z "$BOX_NAME" ] || [ -z "$RESET_JOB_NAME" ] || [ -z "$GET_STATUS_JOB_NAME" ] || [ -z "$APP_ID" ] || [ -z "$BUS_DATE_PARAM" ]; then
  echo "USAGE: $0 <BOX_NAME> <RESET_JOB_NAME> <GET_STATUS_JOB_NAME> <APP_ID> <BUS_DATE_PARAM>" >&2
  exit 97
fi

# --- Paths from env ---
GET_STATUS_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_get_status.ksh
RESET_SH=${AUTOSYS_COMMON_BIN_DIR}/autosys_reset.ksh
SENDEVENT_BIN=${AUTOSYS_RESET_DIR}/sendevent

# Minimal sanity checks
[ -x "$GET_STATUS_SH" ] || { echo "ERROR: missing $GET_STATUS_SH" >&2; exit 91; }
[ -x "$RESET_SH"     ] || { echo "ERROR: missing $RESET_SH"     >&2; exit 92; }
[ -x "$SENDEVENT_BIN" ] || { echo "ERROR: missing $SENDEVENT_BIN" >&2; exit 93; }

# --- Tiny helpers ---
ts()  { date '+%Y-%m-%d %H:%M:%S'; }
log() { echo "`ts` | autosys_controller | $*"; }

# Safe message quoting for ODM logging (handles ' and &)
_esc() {
  _TMP=`echo "$1" | sed "s/'/''/g"`
  _TMP=`echo "$_TMP" | sed "s/&/\\\&/g"`
  echo "$_TMP"
}
MSG_START="'Execution Started autosys_controller for $(_esc "$BOX_NAME")'"
MSG_END="'Execution End autosys_controller for $(_esc "$BOX_NAME")'"

# --- Start log to ODM (non-fatal if helper fails) ---
fx_sam8_odm_event_log "$APP_ID" 1 "$BUS_DATE_PARAM" 93 1 $MSG_START >/dev/null 2>&1
_rc=$?; log "ODM start RC=$_rc"

log "BOX=$BOX_NAME | GET_STATUS_JOB=$GET_STATUS_JOB_NAME | RESET_JOB=$RESET_JOB_NAME | APP_ID=$APP_ID | BUS_DATE=$BUS_DATE_PARAM"
log "BINARIES: GET_STATUS_SH=$GET_STATUS_SH | RESET_SH=$RESET_SH | SENDEVENT_BIN=$SENDEVENT_BIN"

# --- 1) get_status ---
log "Running get_status..."
/bin/ksh "$GET_STATUS_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
RC=$?; log "get_status RC=$RC"
if [ $RC -ne 0 ]; then
  log "get_status failed RC=$RC"
  fx_sam8_odm_event_log "$APP_ID" 1 "$BUS_DATE_PARAM" 94 1 $MSG_END >/dev/null 2>&1
  exit 11
fi
log "get_status OK"

# --- 2) reset (puts box inactive) ---
log "Running reset..."
/bin/ksh "$RESET_SH" "$BOX_NAME" "$APP_ID" "$BUS_DATE_PARAM"
RC=$?; log "reset RC=$RC"
if [ $RC -ne 0 ]; then
  log "reset failed RC=$RC"
  fx_sam8_odm_event_log "$APP_ID" 1 "$BUS_DATE_PARAM" 94 1 $MSG_END >/dev/null 2>&1
  exit 12
fi
log "reset OK"

# --- 3) start the box now ---
log "Starting box via sendevent..."
"$SENDEVENT_BIN" -P 1 -E STARTJOB -J "$BOX_NAME"
RC=$?; log "sendevent RC=$RC"
if [ $RC -ne 0 ]; then
  log "sendevent STARTJOB failed RC=$RC"
  fx_sam8_odm_event_log "$APP_ID" 1 "$BUS_DATE_PARAM" 94 1 $MSG_END >/dev/null 2>&1
  exit 13
fi
log "sendevent OK"

# --- End log to ODM ---
fx_sam8_odm_event_log "$APP_ID" 1 "$BUS_DATE_PARAM" 94 1 $MSG_END >/dev/null 2>&1
_rc=$?; log "ODM end RC=$_rc"

log "DONE."
exit 0