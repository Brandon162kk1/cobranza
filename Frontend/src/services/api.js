const API_SUNAT_URL = import.meta.env.VITE_API_SUNAT;
const API_RENIEC_URL = import.meta.env.VITE_API_RENIEC;

const API_KEY_SUNAT = import.meta.env.VITE_API_KEY_SUNAT;
const API_KEY_RENIEC = import.meta.env.VITE_API_KEY_RENIEC;

export async function consultarRuc(ruc) {

    const response = await fetch(

        `${API_SUNAT_URL}/consultar-ruc`,

        {

            method: "POST",

            headers: {

                "Content-Type": "application/json",

                "x-api-key": API_KEY_SUNAT

            },

            body: JSON.stringify({

                ruc: ruc

            })

        }

    );

    if (!response.ok) {

        throw new Error();

    }

    return await response.json();

}

export async function consultarDni(dni) {
    const response = await fetch(
        `${API_RENIEC_URL}/consultar-dni`,
        {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "x-api-key": API_KEY_RENIEC
            },
            body: JSON.stringify({
                dni: dni
            })
        }
    );

    if (!response.ok) {
        throw new Error();
    }

    return await response.json();
}