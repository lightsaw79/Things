#!/bin/ksh

# $Header: $

# $DateTime: $

# $Change: $

# $Author: $

export ENV_FILE=/apps/samd/activate/package_utilities/common/config/sam8.env

if [ ! -f $ENV_FILE ]
then
echo “Environment file [$ENV_FILE] not found. Exiting …”
exit -1
fi

. $ENV_FILE

box_name=$1
APP_ID=$2
BUS_DATE_PARAM=$3
CATEGORY_ID=1

SCRIPT_NAME=`basename $0 | cut -d'.' -f1`
TIMESTAMP=`date '+%Y%m%d%H%M'`

if [ $APP_ID -eq 15362 ] || [ $APP_ID -eq 15304 ] # $App_Id’s for IRIS jobs force start
then
DAILY_SCHEDULE_BUSS_DATE=`date +'%Y%m%d'`
LCS_BUS_DATE=${BUS_DATE_PARAM:-$DAILY_SCHEDULE_BUSS_DATE}
else
LCS_BUS_DATE=${BUS_DATE_PARAM:-$SAM8_BUSINESS_DATE}
fi

#————————————————————

# Job Execution Log Call

#————————————————————
fx_sam8_oda_event_log $APP_ID $CATEGORY_ID $LCS_BUS_DATE 93 1 “‘Execution Started $SCRIPT_NAME and $APP_ID’”
LogError $APP_ID $? 1 in oda Start Event Log for ‘$APP_ID’ “

############################################

#### AUTOSYS JOB STATUS CHECK START - 1

############################################

# This code block need to be used in conjunction with code block “AUTOSYS JOB STATUS CHECK START - 2”

trap “TRAP_FAILURE_EXE_INDC ERR 1 2 3 4 5 SIGKUP SIGINT SIGTERM

CHKJODBATCHSCHEDULESTART -$APP_ID” -$CATEGORY_ID” -$LCS_BUS_DATE” -$CALENDAR_CD” -$PRE_RUN_OVERRIDE” -$PRE_RUN_MSG” -$GLBL_OVERRIDE” -$GLBL_RE_RUN_MSG”

############################################

#### AUTOSYS JOB STATUS CHECK ENDS - 1

############################################

SAM8_BOX_AUTOSYS_LOG_FILE=$MODEL_BATCH_LOGFILES_DIR/${box_name}_get_status.$SAM8_BUSINESS_DATE.$TIMESTAMP.$$.txt
SAM8_BOX_AUTOSYS_PARSE_FILE=$MODEL_BATCH_LOGFILES_DIR/${box_name}_get_status.$SAM8_BUSINESS_DATE.$TIMESTAMP.$$.txt

$PERL_CMD $HUB_BIN_DIR/ParseAutosysJobs.pl $box_name $SAM8_BOX_AUTOSYS_LOG_FILE $SAM8_BOX_AUTOSYS_PARSE_FILE
LogError 0 $? “Parsing the autosys job log for $box_name”
export TIMEOUT=1800

Log “Log file: $SAM8_BOX_AUTOSYS_PARSE_FILE.err.log”

$AUTOSYS_RESET_DIR/jil < $SAM8_BOX_AUTOSYS_PARSE_FILE > $SAM8_BOX_AUTOSYS_PARSE_FILE.err.log 2>&1
LogError 0 $? “Force Starting All Jobs for $box_name”

Log “Log file: $MODEL_BATCH_LOGFILES_DIR/${box_name}_force_start.log”

$AUTOSYS_RESET_DIR/sendevent -F 1 -E FORCE_STARTJOB -J ${box_name} > $MODEL_BATCH_LOGFILES_DIR/${box_name}_force_start.log 2>&1
LogError 0 $? “Sent Force Start command for $box_name”

############################################

#### AUTOSYS JOB STATUS CHECK START - 2

############################################

# This code block need to be used in conjunction with code block “AUTOSYS JOB STATUS CHECK START - 1”

CHKJODBATCHSCHEDULEEND $APP_ID $LCS_BUS_DATE

############################################

#### AUTOSYS JOB STATUS CHECK ENDS - 2

############################################

fx_sam8_oda_event_log $APP_ID $CATEGORY_ID $LCS_BUS_DATE 94 1 “‘Execution End: $SCRIPT_NAME and $APP_ID’”
LogError $APP_ID $? 1 in oda End Event Log for ‘$APP_ID’ “