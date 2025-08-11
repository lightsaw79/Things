#!/bin/ksh
# controller_inline.ksh
# Usage: controller_inline.ksh <BOX_NAME>

set -u
BOX="$1"
#!/bin/ksh
# controller_inline.ksh
# Usage: controller_inline.ksh <BOX_NAME>

BOX="$1"

# --- get status using your org's wrapper ---
if [ -x "${AUTOSYS_RESET_DIR}/autorep.sh" ]; then
  STATUS=$("${AUTOSYS_RESET_DIR}/autorep.sh" -q "$BOX" 2>/dev/null | awk 'NR==2{print $3}')
else
  STATUS=$(autorep -J "$BOX" -q 2>/dev/null | awk 'NR==2{print $3}')
fi

case "$STATUS" in
  RUNNING|ACTIVATED)
    exit 0
    ;;
  SUCCESS)
    # --- BUSINESS DATE PLACEHOLDER ---
    # If rolling to next business date, uncomment:
    # sendevent -E CHANGE_STATUS -s INACTIVE -J "$BOX" >/dev/null 2>&1
    # sendevent -E FORCE_STARTJOB              -J "$BOX" >/dev/null 2>&1
    exit 0
    ;;
  *)  # FAILURE / TERMINATED / INACTIVE / ON_HOLD / ON_ICE / UNKNOWN
    sendevent -E CHANGE_STATUS -s INACTIVE -J "$BOX" >/dev/null 2>&1
    sendevent -E FORCE_STARTJOB              -J "$BOX" >/dev/null 2>&1
    exit $?
    ;;
esac







# status = 3rd column, 2nd line
STATUS=$(autorep -J "$BOX" -q 2>/dev/null | awk 'NR==2{print $3}')

case "$STATUS" in
  RU*|ST*|AC*|QW*)             # running/starting/queued/activated
    exit 0                     # do nothing; next hourly run will re-check
    ;;
  SU*)                         # success
    # --- BUSINESS DATE PLACEHOLDER ---
    # TODO: if more business dates remain, uncomment the two lines below:
    # sendevent -E CHANGE_STATUS -s INACTIVE -J "$BOX" >/dev/null 2>&1
    # sendevent -E FORCE_STARTJOB              -J "$BOX" >/dev/null 2>&1
    exit 0
    ;;
  *)                           # failure/terminated/on_hold/on_ice/inactive/unknown
    # --- INLINE "my force-start" steps (no external script) ---
    sendevent -E CHANGE_STATUS -s INACTIVE -J "$BOX" >/dev/null 2>&1
    sendevent -E FORCE_STARTJOB              -J "$BOX" >/dev/null 2>&1
    exit $?
    ;;
esac