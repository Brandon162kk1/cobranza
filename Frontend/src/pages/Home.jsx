import { useNavigate } from "react-router-dom";

function Home() {
    const navigate = useNavigate();

    return (
        <div className="container">
            <h1>Jishu Services</h1>
            <h3>Portal de Consultas Rápidas</h3>

            <div className="cards">
                <div
                    className="card-button"
                    onClick={() => navigate("/sunat")}
                >
                    <div className="icon-wrapper">
                        <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path>
                            <polyline points="14 2 14 8 20 8"></polyline>
                            <line x1="16" y1="13" x2="8" y2="13"></line>
                            <line x1="16" y1="17" x2="8" y2="17"></line>
                            <line x1="10" y1="9" x2="8" y2="9"></line>
                        </svg>
                    </div>
                    <h2>Consulta SUNAT</h2>
                    <p>Consulta información de contribuyentes mediante RUC directamente desde el padrón de SUNAT.</p>
                </div>

                <div
                    className="card-button"
                    onClick={() => navigate("/reniec")}
                >
                    <div className="icon-wrapper">
                        <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                            <circle cx="9" cy="7" r="4"></circle>
                            <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                            <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                        </svg>
                    </div>
                    <h2>Consulta RENIEC</h2>
                    <p>Valida la identidad de ciudadanos peruanos consultando el número de DNI.</p>
                </div>
            </div>
        </div>
    );
}

export default Home;