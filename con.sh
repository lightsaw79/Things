#!/bin/ksh
# Combined: get status -> reset -> start (force/start)
# Uses your org’s exact logic as in autosys_get_status.ksh and autosys_reset.ksh

###############################################################################
# Common header + env (identical to both scripts)
###############################################################################
# $Header: $
# $DateTime: $
# $Change: $
# $Author: $

export ENV_FILE=/apps/sam8/actimize/package_utilities/common/config/sam8.env

if [ ! -f $ENV_FILE ]
then
  echo "Environment file [$ENV_FILE] not found. Exiting ..."
  exit -1
fi

. $ENV_FILE

# --- ARGS (same order you use) ---
box_name=$1
res=$2
get_s=$3
APP_ID=$4
BUS_DATE_PARAM=$5
CATEGORY_ID=1

SCRIPT_NAME=`basename $0 | cut -d"." -f1`
TIMESTAMP=`date +"%Y%m%d%H%M"`

# Business date logic (kept identical from your screenshots)
# (Your get_status image shows 15361/15363; reset shows 15362/15364.
# If that’s intentional, keep it; otherwise add/remove IDs below.)
if [ $APP_ID -eq 15361 ] || [ $APP_ID -eq 15362 ] || [ $APP_ID -eq 15363 ] || [ $APP_ID -eq 15364 ]
then
  DAILY_SCHEDULE_BUSS_DATE=`date +"%Y%m%d"`
  LCS_BUS_DATE=${BUS_DATE_PARAM:-$DAILY_SCHEDULE_BUSS_DATE}
else
  LCS_BUS_DATE=${BUS_DATE_PARAM:-$SAM8_BUSINESS_DATE}
fi

###############################################################################
# Job Execution Log Call (start)  — appears in both scripts
###############################################################################
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $LCS_BUS_DATE 93 1 "Execution Started $SCRIPT_NAME and $APP_ID"
LogError $APP_ID $? "In odm Start Event Log for '$APP_ID' "

###############################################################################
# =============  GET STATUS SECTION (exact flow from autosys_get_status.ksh) ==
###############################################################################

##### AUTOSYS JOB STATUS CHECK START - 1 #####
# This code block need to be used in conjunction with code block "AUTOSYS JOB STATUS CHECK START - 2"
trap "TRAP_FAILURE_EXE_LOG" ERR 2 3 4 5 SIGHUP SIGINT SIGTERM

CHKOMDBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$LCS_BUS_DATE" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

##### AUTOSYS JOB STATUS CHECK ENDS - 1 #####

# --- your get-status middle block ---
Log "Executing autorep command"
Log "Log file: $MODEL_BATCH_LOGFILES_DIR/${box_name}.get_status.log"
$AUTOSYS_RESET_DIR/autorep.sh -q "$box_name" > "$MODEL_BATCH_LOGFILES_DIR/${box_name}.get_status.log" 2>&1
LogError 0 $? "autorep command execution"

##### AUTOSYS JOB STATUS CHECK START - 2 #####
# This code block need to be used in conjunction with code block "AUTOSYS JOB STATUS CHECK START - 1"
CHKOMDBATCHSCHEDULEEND $APP_ID $LCS_BUS_DATE
##### AUTOSYS JOB STATUS CHECK ENDS - 2 #####

###############################################################################
# ====================  RESET SECTION (exact flow from reset.ksh)  ============
###############################################################################

##### AUTOSYS JOB STATUS CHECK START - 1 #####
# This code block need to be used in conjunction with code block "AUTOSYS JOB STATUS CHECK START - 2"
trap "TRAP_FAILURE_EXE_LOG" ERR 2 3 4 5 SIGHUP SIGINT SIGTERM

CHKOMDBATCHSCHEDULESTART "$APP_ID" "$CATEGORY_ID" "$LCS_BUS_DATE" "$CALNDR_CD" "$RE_RUN_OVERRIDE" "$RE_RUN_MSG" "$GLBL_OVERRIDE" "$GLBL_RE_RUN_MSG"

##### AUTOSYS JOB STATUS CHECK ENDS - 1 #####

# exact variables from your reset screenshots
SAMB_BOX_AUTOSYS_LOG_FILE="$MODEL_BATCH_LOGFILES_DIR/${box_name}.get_status.$SAM8_BUSINESS_DATE.$TIMESTAMP.$$.txt"
SAMB_BOX_AUTOSYS_PARSE_FILE="$MODEL_BATCH_LOGFILES_DIR/${box_name}.get_status.$SAM8_BUSINESS_DATE.$TIMESTAMP.$$.parse"

# Parse jobs
$PERL_CMD $HUB_BIN_DIR/ParseAutosysJobs.pl $box_name $SAMB_BOX_AUTOSYS_LOG_FILE $SAMB_BOX_AUTOSYS_PARSE_FILE
LogError 0 $? "Parsing the autosys job log for $box_name"
export TIMEOUT=1800

# Apply JIL updates (exact)
Log "Log file: ${SAMB_BOX_AUTOSYS_PARSE_FILE}.err.log"
$AUTOSYS_RESET_DIR/jil < $SAMB_BOX_AUTOSYS_PARSE_FILE > ${SAMB_BOX_AUTOSYS_PARSE_FILE}.err.log 2>&1
LogError 0 $? "Updating All jobs for $box_name"

# Send INACTIVE (exact)
Log "Log file: $MODEL_BATCH_LOGFILES_DIR/${box_name}_inactive.log"
$AUTOSYS_RESET_DIR/sendevent -p 1 -E CHANGE_STATUS -s INACTIVE -J ${box_name} > $MODEL_BATCH_LOGFILES_DIR/${box_name}_inactive.log 2>&1
LogError 0 $? "Sent Inactive command for $box_name"

##### AUTOSYS JOB STATUS CHECK START - 2 #####
# This code block need to be used in conjunction with code block "AUTOSYS JOB STATUS CHECK START - 1"
CHKOMDBATCHSCHEDULEEND $APP_ID $LCS_BUS_DATE
##### AUTOSYS JOB STATUS CHECK ENDS - 2 #####

###############################################################################
# ======================  START SECTION (new, tiny addition)  =================
###############################################################################
# Choose: START_MODE=start  -> respects conditions (STARTJOB)
#         START_MODE=force  -> ignores conditions (FORCE_STARTJOB)  [default]
: ${START_MODE:=force}
if [ "$START_MODE" = "start" ]; then
  _EVENT="STARTJOB"
else
  _EVENT="FORCE_STARTJOB"
fi

Log "Starting box with $_EVENT: $box_name"
START_LOG="$MODEL_BATCH_LOGFILES_DIR/${box_name}.start.$SAM8_BUSINESS_DATE.$TIMESTAMP.$$.log"
$AUTOSYS_RESET_DIR/sendevent -p 1 -E $_EVENT -J "$box_name" > "$START_LOG" 2>&1
LogError 0 $? "$_EVENT for $box_name (see $START_LOG)"

###############################################################################
# Job Execution Log Call (end)
###############################################################################
fx_sam8_odm_event_log $APP_ID $CATEGORY_ID $LCS_BUS_DATE 94 1 "Execution End: $SCRIPT_NAME and $APP_ID"
LogError $APP_ID $? "In odm End Event Log for '$APP_ID' "