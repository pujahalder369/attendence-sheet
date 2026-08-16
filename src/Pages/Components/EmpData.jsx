
const EmpData = ({ data }) => {
    return (
        <div>
            {data.results.map((emp) => {
                let present = 0;
                let absent = 0;
                let holiday = 0;
                let weekoff = 0;
                let halfday = 0;

                emp.attendances.forEach((a) => {
                    if (a.absent) {
                        absent++;
                    } else if (a.holiday) {
                        holiday++;
                    } else if (a.week_off) {
                        weekoff++;
                    } else if (a.half_day) {
                        halfday++;
                    } else if (a.leave) {
                        // Leave → don't count as present
                    } else {
                        present++;
                    }
                });

                return (
                    <div
                        key={emp?.employee?.id}
                        className='border-2 rounded-lg p-5 mb-5'>
                        <div className='flex justify-between mb-5 text-lg'>
                            <p className='font-semibold'>{emp?.employee?.first_name} {emp?.employee?.first_name}</p>
                            <span className='font-semibold'>{emp?.employee?.department_name || "---"}</span>
                        </div>
                        <div className='flex justify-between gap-2 text-center'>
                            <div>
                                <p>Present</p>
                                <span>{present}</span>
                            </div>
                            <div>
                                <p>Absent</p>
                                <span>{absent}</span>
                            </div>
                            <div>
                                <p>Halfday</p>
                                <span>{halfday}</span>
                            </div>
                            <div>
                                <p>Weekoff</p>
                                <span>{weekoff}</span>
                            </div>
                            <div>
                                <p>Holiday</p>
                                <span>{holiday}</span>
                            </div>
                        </div>
                    </div>
                )
            })}
        </div>
    )
}

export default EmpData
