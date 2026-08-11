import { useState } from 'react'
import PDFPopup from '../../Components/pdfPopup/PDFPopup';

const RightPart = () => {
    const [pdfPopup, setPdfPopup] = useState(false);

    return (
        <div>
            <div className="bg-sky-400 mx-5 p-5 rounded-lg h-full w-[92%] md:w-[95%] lg:w-[280px] xl:w-[350px]">
                <div className="flex justify-center align-center font-semibold text-xl bg-black rounded text-white capitalize py-3 mb-10">
                    Download Report
                </div>

                <div className="flex flex-col justify-center items-center gap-3 lg:gap-5">
                    <button
                        onClick={() => setPdfPopup(true)}
                        type="button"
                        className="bg-[#EDEB62] p-3 px-12 text-xl rounded-lg inline-flex lg:w-full justify-center font-semibold"
                    >
                        Report 1
                    </button>
                    {pdfPopup && <PDFPopup closePopup={() => setPdfPopup(false)} />}
                </div>
            </div>
        </div>
    )
}

export default RightPart
