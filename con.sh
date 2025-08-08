#!/bin/ksh

ENV_FILE=/apps/samd/actimize/package_utilities/common/config/sam8.env
if [ ! -f "$ENV_FILE" ]; then
  print "Environment file [$ENV_FILE] not found. Exiting ..."
  exit -1
fi
. "$ENV_FILE"

BOX_NAME=$1
GET_STATUS_JOB=$2
RESET_JOB=$3
APP_ID=$4
BUS_DATE_PARAM=$5
CATEGORY_ID=1


SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
TIMESTAMP=`date +"%Y%m%d%H%M%S"`

LOG_FILE="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.log"
ERR_FILE="${MODEL_BATCH_LOGFILES_DIR}/${BOX_NAME}_ctrl_${TIMESTAMP}.err"

AUTOREP_CMD="${AUTOSYS_RESET_DIR}/autorep.sh"
SENDEVENT_CMD="${AUTOSYS_RESET_DIR}/sendevent"

log() { print -- "`date '+%Y-%m-%d %H:%M:%S'` | $SCRIPT_NAME | $*"; }
logf() { log "$*" | tee -a "$LOG_FILE"; }
logerrf() { log "ERROR: $*" | tee -a "$ERR_FILE" 1>&2; }

get_status() {
    $AUTOREP_CMD -q "$1" 2>/dev/null | awk 'NR==1{print $2}'
}

wait_for_end() {
    JOB=$1
    TIMEOUT=$2
    i=0
    while :; do
        STATUS=`get_status "$JOB"`
        case "$STATUS" in
            SU|FA|TE) logf "$JOB completed with status $STATUS"; return 0 ;;
        esac
        [ $i -ge $TIMEOUT ] && { logerrf "Timeout waiting for $JOB"; return 1; }
        i=`expr $i + 1`
        sleep 5
    done
}

wait_for_inactive() {
    JOB=$1
    TIMEOUT=$2
    i=0
    while :; do
        STATUS=`get_status "$JOB"`
        [ "$STATUS" = "INACTIVE" ] && { logf "$JOB is INACTIVE"; return 0; }
        [ $i -ge $TIMEOUT ] && { logerrf "Timeout waiting for $JOB to be INACTIVE"; return 1; }
        i=`expr $i + 1`
        sleep 5
    done
}

fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 93 1 "Execution Started $SCRIPT_NAME for $BOX_NAME"

CHKDOMBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$BUS_DATE_PARAM" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

CUR_STATUS=`get_status "$BOX_NAME"`
logf "Current status of $BOX_NAME = $CUR_STATUS"

if [ "$CUR_STATUS" != "SU" ]; then
    logf "Box not SU → FORCE_START $GET_STATUS_JOB"
    $SENDEVENT_CMD -E FORCE_STARTJOB -J "$GET_STATUS_JOB" >>"$LOG_FILE" 2>>"$ERR_FILE"
    wait_for_end "$GET_STATUS_JOB" 120

    logf "FORCE_START $RESET_JOB"
    $SENDEVENT_CMD -E FORCE_STARTJOB -J "$RESET_JOB" >>"$LOG_FILE" 2>>"$ERR_FILE"
    wait_for_end "$RESET_JOB" 240
    wait_for_inactive "$BOX_NAME" 240
else
    logf "Box is SU → No reset needed"
fi

logf "FORCE_START $BOX_NAME for next cycle"
$SENDEVENT_CMD -E FORCE_STARTJOB -J "$BOX_NAME" >>"$LOG_FILE" 2>>"$ERR_FILE"

CHKDOMBATCHSCHEDULEEND "$APP_ID" "$BUS_DATE_PARAM"

fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $BUS_DATE_PARAM 94 1 "Execution End $SCRIPT_NAME for $BOX_NAME"
exit 0