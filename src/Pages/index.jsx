import { useEffect, useState } from "react";
import LeftPart from './Components/LeftPart';
import RightPart from './Components/RightPart';
import EmpData from './Components/EmpData';

const AttendenceSheet = () => {
    const [data, setData] = useState({ results: [] });

    useEffect(() => {
        fetch("/data.json")
            .then((res) => res.json())
            .then((json) => setData(json))
            .catch((err) => console.log(err));
    }, []);

    return (
        <>
            <div className="flex flex-col lg:flex-row gap-5 lg:gap-0 bg-[#EDEB62] h-auto lg:h-[100vh] w-full py-5">
                <LeftPart />

                <div className="bg-white flex-1 h-full w-full p-5 rounded-lg overflow-y-scroll custom-scrollbar">
                    <EmpData data={data} />
                </div>

                <RightPart />
            </div>
        </>
    )
}

export default AttendenceSheet;
