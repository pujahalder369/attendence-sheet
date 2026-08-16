export const formatDuration = (seconds = 0) => {
  const totalSeconds = Number(seconds) || 0;

  if (totalSeconds <= 0) {
    return "-";
  }

  const hrs = Math.floor(totalSeconds / 3600);
  const mins = Math.floor((totalSeconds % 3600) / 60);

  return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
};

export const getStatus = (attendance) => {
  if (!attendance) return "-";
  if (attendance.absent) return "A";
  if (attendance.leave) return "L";
  if (attendance.half_day) return "HD";
  if (attendance.holiday) return "HO";
  if (attendance.week_off) return "WO";

  return "P";
};

export const calculateSummary = (attendances = []) => {
  let present = 0;
  let absent = 0;
  let leave = 0;
  let holiday = 0;
  let weekoff = 0;
  let halfday = 0;

  let totalWorkingSeconds = 0;
  let totalOTSeconds = 0;

  attendances.forEach((item) => {
    if (!item) return;
    if (item.absent) {
      absent++;
      return;
    }
    if (item.leave) {
      leave++;
      return;
    }
    if (item.week_off) {
      weekoff++;
      return;
    }
    if (item.holiday) {
      holiday++;
      return;
    }
    if (item.half_day) {
      halfday++;
      present++;
    } else {
      present++;
    }

    totalWorkingSeconds += Number(item.duration) || 0;
    totalOTSeconds += Number(item.ot) || 0;
  });

  return [
    {
      label: "Present",
      value: present,
    },
    {
      label: "Holiday",
      value: holiday,
    },
    {
      label: "WeekOff",
      value: weekoff,
    },
    {
      label: "Halfday",
      value: halfday,
    },
    {
      label: "Absent",
      value: absent,
    },
    {
      label: "Leave",
      value: leave,
    },
    {
      label: "Work Hrs",
      value: formatDuration(totalWorkingSeconds),
    },
    {
      label: "OT Hrs",
      value: formatDuration(totalOTSeconds),
    },
  ];
};