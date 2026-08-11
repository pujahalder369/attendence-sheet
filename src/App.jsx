import { BrowserRouter } from "react-router-dom";
import "./App.css";
import ToasterWrapper from "./Components/toastWrapper/ToasterWrapper";
import MainRoutes from "./Routes";

function App() {
  return (
    <>
      <BrowserRouter>
        <MainRoutes />
        <ToasterWrapper />
      </BrowserRouter>
    </>
  );
}

export default App;
