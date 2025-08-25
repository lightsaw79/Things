get_status() {
  "${AUTOSYS_RESET_DIR}/autorep.sh" -q "$box_name" 2>/dev/null \
  | sed -n '2p' \
  | tr -d '\r' \
  | grep -Eo 'RUNNING|ACTIVATED|SUCCESS|FAILURE|TERMINATED|ON_HOLD|ON_ICE|INACTIVE' \
  | head -1
}

# ---------- Inline force-start (your exact pattern) ----------
force_start_now() {
  "${AUTOSYS_RESET_DIR}/sendevent" -p 1 -E CHANGE_STATUS -s INACTIVE -J "$box_name" \
    > "${logdir}/${box_name}_inactive.log" 2>&1
  "${AUTOSYS_RESET_DIR}/sendevent" -E FORCE_STARTJOB -J "$box_name" \
    > "${logdir}/${box_name}.log" 2>&1
}

# ---------- Wait helpers ----------
wait_until_success() {
  # Wait until box reaches SUCCESS; if it hits a non-running/non-success state, restart.
  while : ; do
    S="$(get_status)"
    case "$S" in
      RUNNING|ACTIVATED)
        sleep 30
        ;;
      SUCCESS)
        return 0
        ;;
      FAILURE|TERMINATED|INACTIVE|ON_HOLD|ON_ICE|'')
        force_start_now
        sleep 5
        ;;
      *)
        force_start_now
        sleep 5
        ;;
    esac
  done
}

# ---------- Drain all pending dates in THIS run ----------
# 1) If currently running, wait for it to finish (become SUCCESS)
while : ; do
  S0="$(get_status)"
  case "$S0" in
    RUNNING|ACTIVATED) sleep 30 ;;
    SUCCESS|FAILURE|TERMINATED|INACTIVE|ON_HOLD|ON_ICE|*) break ;;
  esac
done

# 2) For each pending business date:
for _d in $BDATES ; do
  # Kick the next run
  force_start_now
  # Wait until this run succeeds
  wait_until_success
done

# 3) All dates drained in this single controller run
exit 0