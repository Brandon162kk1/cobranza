import { BrowserRouter, Routes, Route } from "react-router-dom";

import Home from "./pages/Home";
import Sunat from "./pages/Sunat";
import Reniec from "./pages/Reniec";

import "./App.css";

function App() {
    return (
        <BrowserRouter>

            <Routes>

                <Route path="/" element={<Home />} />

                <Route path="/sunat" element={<Sunat />} />

                <Route path="/reniec" element={<Reniec />} />

            </Routes>

        </BrowserRouter>
    );
}

export default App;