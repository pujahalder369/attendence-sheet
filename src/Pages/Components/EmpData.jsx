
const EmpData = ({ data }) => {
    return (
        <div>
            {data.results.map((emp) => {
                const present = emp.attendances.filter((a) =>
                    !a.absent &&
                    !a.holiday &&
                    !a.week_off &&
                    !a.half_day &&
                    !a.leave
                ).length;
                const absent = emp.attendances.filter((a) => a.absent).length;
                const holiday = emp.attendances.filter((a) => a.holiday).length;
                const weekoff = emp.attendances.filter((a) => a.week_off).length;
                const halfday = emp.attendances.filter((a) => a.half_day).length;

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
