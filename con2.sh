#!/bin/ksh
# controller_inline_drain.ksh
# Usage:
#   controller_inline_drain.ksh <BOX_NAME> [<DATE_LIST>]
# Example:
#   controller_inline_drain.ksh TML_DUMMY_TEST_BOX "2025-03-11,2025-03-12,2025-03-14,2025-03-14"
#
# Notes:
# - <DATE_LIST> is a placeholder list; we are NOT passing dates to Autosys.
#   We just use the list to decide how many times to inactivate + force-start.
# - Uses your org paths if set: $AUTOSYS_RESET_DIR and $MODEL_BATCH_LOGFILES_DIR.

box_name="$1"
BDATES="${2:-${BDATES:-}}"

# Default log dir if your env var isn't set
LOGDIR="${MODEL_BATCH_LOGFILES_DIR:-/tmp}"

# --- helpers (lean) ---
get_status() {
  if [ -x "${AUTOSYS_RESET_DIR}/autorep.sh" ]; then
    "${AUTOSYS_RESET_DIR}/autorep.sh" -q "$box_name" 2>/dev/null | awk 'NR==2{print $3}'
  else
    autorep -J "$box_name" -q 2>/dev/null | awk 'NR==2{print $3}'
  fi
}

consume_next_date() {
  # Drop first token (comma/space separated)
  BDATES=$(echo "${BDATES}" | sed 's/^[^ ,]*[ ,]*//')
}

# Inline "my force-start": INACTIVE -> FORCE_STARTJOB (your exact pattern)
force_start_now() {
  # Make INACTIVE
  "${AUTOSYS_RESET_DIR}/sendevent" -p 1 -E CHANGE_STATUS -s INACTIVE -J "${box_name}" \
    > "${LOGDIR}/${box_name}_inactive.log" 2>&1
  # FORCE START
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "${box_name}" \
    > "${LOGDIR}/${box_name}.log" 2>&1
}

# ---------------- MAIN (drain remaining dates) ----------------
# Loop until date list is empty. Between runs, wait for RUNNING to finish.
while : ; do
  STATUS="$(get_status)"

  case "$STATUS" in
    RUNNING|ACTIVATED)
      sleep 60
      continue
      ;;

    SUCCESS)
      # If there are remaining dates, trigger next run; else finish.
      if [ -n "${BDATES}" ]; then
        force_start_now
        consume_next_date
        sleep 5
        continue
      else
        # All requested dates covered; mark controller success.
        exit 0
      fi
      ;;

    FAILURE|TERMINATED|INACTIVE|ON_HOLD|ON_ICE|'')
      # Not running and not success -> kick it off for the current/next date
      force_start_now
      sleep 5
      continue
      ;;

    *)
      # Any other unexpected state -> same recovery
      force_start_now
      sleep 5
      continue
      ;;
  esac
done