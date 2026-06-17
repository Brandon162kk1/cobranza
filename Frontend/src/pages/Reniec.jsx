import { useState } from "react";
import { Link } from "react-router-dom";
import { consultarDni } from "../services/api";

function Reniec() {
    const [dni, setDni] = useState("");
    const [resultado, setResultado] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState("");
    const [copiado, setCopiado] = useState(false);

    const buscar = async () => {
        setError("");
        setResultado(null);

        if (dni.length !== 8 || !/^\d+$/.test(dni)) {
            setError("Ingrese un DNI válido de 8 dígitos numéricos.");
            return;
        }

        try {
            setLoading(true);
            const data = await consultarDni(dni);
            setResultado(data);
        } catch {
            setError("Error al consultar RENIEC. Verifique el número de DNI.");
        } finally {
            setLoading(false);
        }
    };

    const limpiar = () => {
        setDni("");
        setResultado(null);
        setError("");
    };

    const copiar = () => {
        if (!resultado) return;

        const texto = `DNI: ${dni}
Nombres: ${resultado.nombres}
Apellido Paterno: ${resultado.apellido_paterno}
Apellido Materno: ${resultado.apellido_materno}
Nombre Completo: ${resultado.nombres} ${resultado.apellido_paterno} ${resultado.apellido_materno}`;

        navigator.clipboard.writeText(texto);
        setCopiado(true);
        setTimeout(() => setCopiado(false), 2000);
    };

    return (
        <div className="container">
            <h1>👤 Consulta RENIEC</h1>
            <h3>Validación rápida de identidad mediante DNI</h3>

            <div className="search-box">
                <div className="input-container">
                    <input
                        type="text"
                        placeholder="Ingrese DNI (8 dígitos)"
                        value={dni}
                        maxLength={8}
                        onChange={(e) => setDni(e.target.value.replace(/\D/g, ""))}
                        onKeyDown={(e) => e.key === "Enter" && buscar()}
                    />
                </div>
                <div className="button-group">
                    <button className="primary" onClick={buscar} disabled={loading}>
                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                            <circle cx="11" cy="11" r="8"></circle>
                            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                        </svg>
                        {loading ? "Buscando..." : "Consultar"}
                    </button>
                    <button className="secondary" onClick={limpiar} disabled={loading}>
                        Limpiar
                    </button>
                </div>
            </div>

            <div style={{ margin: "24px 0" }}>
                <Link to="/" className="back-link">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="19" y1="12" x2="5" y2="12"></line>
                        <polyline points="12 19 5 12 12 5"></polyline>
                    </svg>
                    Volver al inicio
                </Link>
            </div>

            {loading && (
                <div className="loader-container">
                    <div className="spinner"></div>
                    <div className="loader-text">Consultando padrón de RENIEC...</div>
                </div>
            )}

            {error && (
                <div className="error-banner">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <circle cx="12" cy="12" r="10"></circle>
                        <line x1="12" y1="8" x2="12" y2="12"></line>
                        <line x1="12" y1="16" x2="12.01" y2="16"></line>
                    </svg>
                    {error}
                </div>
            )}

            {resultado && (
                <div className="result-card">
                    <div className="result-header">
                        <h2 className="result-title">
                            <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: "#3b82f6" }}>
                                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                                <circle cx="12" cy="7" r="4"></circle>
                            </svg>
                            Ficha DNI
                        </h2>
                        <span className="badge badge-active">
                            Encontrado
                        </span>
                    </div>

                    <div style={{ display: "flex", gap: "24px", alignItems: "stretch", flexWrap: "wrap", marginBottom: "24px" }}>
                        <div className="result-section" style={{ flex: 1, minWidth: "280px", margin: 0 }}>
                            <h3 className="section-heading">
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                    <circle cx="12" cy="12" r="10"></circle>
                                    <line x1="12" y1="16" x2="12" y2="12"></line>
                                    <line x1="12" y1="8" x2="12.01" y2="8"></line>
                                </svg>
                                Datos de Identidad
                            </h3>
                            <div className="data-row">
                                <span className="data-label">Número de DNI</span>
                                <span className="data-value" style={{ letterSpacing: "1px" }}>{dni}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Nombres</span>
                                <span className="data-value">{resultado.nombres}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Apellido Paterno</span>
                                <span className="data-value">{resultado.apellido_paterno}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Apellido Materno</span>
                                <span className="data-value">{resultado.apellido_materno}</span>
                            </div>
                            <div className="data-row" style={{ borderTop: "1.5px solid #e2e8f0", marginTop: "8px", paddingTop: "12px" }}>
                                <span className="data-label" style={{ color: "#0f172a", fontWeight: "700" }}>Nombre Completo</span>
                                <span className="data-value" style={{ color: "#2563eb", fontWeight: "700" }}>
                                    {resultado.nombres} {resultado.apellido_paterno} {resultado.apellido_materno}
                                </span>
                            </div>
                        </div>

                        <div style={{
                            width: "140px",
                            minHeight: "175px",
                            background: "#f8fafc",
                            border: "1.5px solid #e2e8f0",
                            borderRadius: "16px",
                            display: "flex",
                            flexDirection: "column",
                            alignItems: "center",
                            justifyContent: "center",
                            padding: "10px",
                            boxSizing: "border-box",
                            boxShadow: "0 4px 6px -1px rgba(0, 0, 0, 0.05)",
                            alignSelf: "center"
                        }}>
                            <div style={{
                                width: "100%",
                                height: "100%",
                                minHeight: "150px",
                                background: "#f1f5f9",
                                borderRadius: "10px",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                color: "#94a3b8",
                                border: "1px solid #e2e8f0"
                            }}>
                                <svg xmlns="http://www.w3.org/2000/svg" width="90" height="90" viewBox="0 0 24 24" fill="currentColor">
                                    <path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/>
                                </svg>
                            </div>
                        </div>
                    </div>

                    <div style={{ display: "flex", justifyContent: "flex-end" }}>
                        <button className="primary" onClick={copiar} style={{ minWidth: "140px" }}>
                            {copiado ? (
                                <>
                                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                                        <polyline points="20 6 9 17 4 12"></polyline>
                                    </svg>
                                    ¡Copiado!
                                </>
                            ) : (
                                <>
                                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                                    </svg>
                                    Copiar Datos
                                </>
                            )}
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}

export default Reniec;
