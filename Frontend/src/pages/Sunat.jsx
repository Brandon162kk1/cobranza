import { useState } from "react";
import { Link } from "react-router-dom";
import { consultarRuc } from "../services/api";

function Sunat() {
    const [ruc, setRuc] = useState("");
    const [resultado, setResultado] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState("");
    const [copiado, setCopiado] = useState(false);

    const buscar = async () => {
        setError("");
        setResultado(null);

        if (ruc.length !== 11 || !/^\d+$/.test(ruc)) {
            setError("Ingrese un RUC válido de 11 dígitos numéricos.");
            return;
        }

        try {
            setLoading(true);
            const data = await consultarRuc(ruc);
            setResultado(data);
        } catch {
            setError("Error al consultar SUNAT. Verifique el número de RUC.");
        } finally {
            setLoading(false);
        }
    };

    const limpiar = () => {
        setRuc("");
        setResultado(null);
        setError("");
    };

    const renderVal = (val) => {
        if (val === null || val === undefined) return "-";
        const cleaned = val.toString().trim();
        return cleaned === "" ? "-" : cleaned;
    };

    const renderActividad = (code, desc) => {
        const cleanCode = renderVal(code);
        const cleanDesc = renderVal(desc);
        if (cleanCode === "-" && cleanDesc === "-") return "-";
        return `[${cleanCode}] ${cleanDesc}`;
    };

    const copiar = () => {
        if (!resultado) return;

        const texto = `Tipo Documento: ${renderVal(resultado.tipo_documento)}
RUC: ${renderVal(resultado.ruc)}
Número Documento: ${renderVal(resultado.numero_documento)}
Nombres: ${renderVal(resultado.nombres)}
Apellidos: ${renderVal(resultado.apellidos)}
Razón Social: ${renderVal(resultado.razon_social)}
Nombre Comercial: ${renderVal(resultado.nombre_comercial)}
Fecha Inicio: ${renderVal(resultado.fecha_inicio)}
Estado: ${renderVal(resultado.estado)}
Dirección Fiscal: ${renderVal(resultado.domicilio_fiscal)}
Provincia: ${renderVal(resultado.provincia)}
Ciudad: ${renderVal(resultado.ciudad)}
Distrito: ${renderVal(resultado.distrito)}
Actividad Principal: ${renderActividad(resultado.cod_principal, resultado.actividad_principal)}
Actividad Secundaria 1: ${renderActividad(resultado.cod_secundario_1, resultado.actividad_1)}
Actividad Secundaria 2: ${renderActividad(resultado.cod_secundario_2, resultado.actividad_2)}`;

        navigator.clipboard.writeText(texto);
        setCopiado(true);
        setTimeout(() => setCopiado(false), 2000);
    };

    const esActivo = resultado?.estado?.toUpperCase() === "ACTIVO";

    return (
        <div className="container">
            <h1>📄 Consulta SUNAT</h1>
            <h3>Búsqueda rápida de contribuyentes en el padrón nacional</h3>

            <div className="search-box">
                <div className="input-container">
                    <input
                        type="text"
                        placeholder="Ingrese RUC (11 dígitos)"
                        value={ruc}
                        maxLength={11}
                        onChange={(e) => setRuc(e.target.value.replace(/\D/g, ""))}
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
                    <div className="loader-text">Consultando padrón de SUNAT...</div>
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
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                <polyline points="14 2 14 8 20 8"></polyline>
                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                <polyline points="10 9 9 9 8 9"></polyline>
                            </svg>
                            Ficha RUC
                        </h2>
                        <span className={`badge ${esActivo ? "badge-active" : "badge-inactive"}`}>
                            {renderVal(resultado.estado)}
                        </span>
                    </div>

                    <div className="result-grid">
                        <div className="result-section">
                            <h3 className="section-heading">
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                    <circle cx="12" cy="12" r="10"></circle>
                                    <line x1="12" y1="16" x2="12" y2="12"></line>
                                    <line x1="12" y1="8" x2="12.01" y2="8"></line>
                                </svg>
                                Información General
                            </h3>
                            <div className="data-row">
                                <span className="data-label">Tipo Doc. de Socio</span>
                                <span className="data-value">{renderVal(resultado.tipo_documento)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">RUC</span>
                                <span className="data-value">{renderVal(resultado.ruc)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Número Documento</span>
                                <span className="data-value">{renderVal(resultado.numero_documento)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Nombres</span>
                                <span className="data-value">{renderVal(resultado.nombres)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Apellidos</span>
                                <span className="data-value">{renderVal(resultado.apellidos)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Razón Social</span>
                                <span className="data-value">{renderVal(resultado.razon_social)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Nombre Comercial</span>
                                <span className="data-value">{renderVal(resultado.nombre_comercial)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Fecha de Inicio</span>
                                <span className="data-value">{renderVal(resultado.fecha_inicio)}</span>
                            </div>
                        </div>

                        <div className="result-section">
                            <h3 className="section-heading">
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                    <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                                    <circle cx="12" cy="10" r="3"></circle>
                                </svg>
                                Ubicación
                            </h3>
                            <div className="data-row">
                                <span className="data-label">Departamento / Prov.</span>
                                <span className="data-value">{renderVal(resultado.provincia)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Ciudad / Provincia</span>
                                <span className="data-value">{renderVal(resultado.ciudad)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Distrito</span>
                                <span className="data-value">{renderVal(resultado.distrito)}</span>
                            </div>
                            <div className="data-row">
                                <span className="data-label">Dirección Fiscal</span>
                                <span className="data-value">{renderVal(resultado.domicilio_fiscal)}</span>
                            </div>
                        </div>
                    </div>

                    <div className="result-section" style={{ marginBottom: "24px" }}>
                        <h3 className="section-heading">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                <rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect>
                                <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>
                            </svg>
                            Actividades Económicas
                        </h3>
                        <div className="data-row">
                            <span className="data-label">Actividad Principal</span>
                            <span className="data-value">
                                {renderActividad(resultado.cod_principal, resultado.actividad_principal)}
                            </span>
                        </div>
                        <div className="data-row">
                            <span className="data-label">Actividad Secundaria 1</span>
                            <span className="data-value">
                                {renderActividad(resultado.cod_secundario_1, resultado.actividad_1)}
                            </span>
                        </div>
                        <div className="data-row">
                            <span className="data-label">Actividad Secundaria 2</span>
                            <span className="data-value">
                                {renderActividad(resultado.cod_secundario_2, resultado.actividad_2)}
                            </span>
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
                                    Copiar Ficha
                                </>
                            )}
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}

export default Sunat;