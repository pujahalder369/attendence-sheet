import { useEffect, useState } from "react";
import LeftPart from './Components/LeftPart';
import RightPart from './Components/RightPart';
import EmpData from './Components/EmpData';
import { toast } from "react-toastify";

const AttendenceSheet = () => {
    const [data, setData] = useState({ results: [] });

    useEffect(() => {
        const dataFetch = async () => {
            try {
                const res = await fetch("/data.json");
                if (!res.ok) {
                    throw new Error("Failed to fetch data")
                }
                const json = await res.json();
                setData(json);
            } catch (err) {
                console.log(err);
                toast.error("Failed to load attendance data");
            }
        }
        dataFetch();
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
