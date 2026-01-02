const { createClient } = require('@supabase/supabase-js');

// Credenciales obtenidas de contador/.env.local (o config.js)
const supabaseUrl = 'https://orpplvjccokivumdbctw.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9ycHBsdmpjY29raXZ1bWRiY3R3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjUyMjQ5NzYsImV4cCI6MjA4MDgwMDk3Nn0.VXtgtmam2virgn6N2ygPPiJ3rhMv0zEZMSL59vwXgcc';

let supabase = null;

function initSupabase() {
    if (!supabase) {
        supabase = createClient(supabaseUrl, supabaseKey);
    }
    return supabase;
}

// Helper para limpiar
function clean(str) {
    return (str || '').toString().trim();
}

// Obtener contactos (email, whatsapp) de AMBAS tablas
async function getSupabaseContacts() {
    try {
        const client = initSupabase();

        // 1. Fetch tabla contacts
        const p1 = client
            .from('contacts')
            .select('email, whatsapp');

        // 2. Fetch tabla activities
        const p2 = client
            .from('activities')
            .select('email, whatsapp');

        const [res1, res2] = await Promise.all([p1, p2]);

        if (res1.error) console.error('Error fetching contacts:', res1.error);
        if (res2.error) console.error('Error fetching activities:', res2.error);

        const list1 = res1.data || [];
        const list2 = res2.data || [];

        // 3. Unificar y dediplicar
        const all = [...list1, ...list2];
        const uniqueMap = new Map();

        all.forEach(item => {
            const e = clean(item.email).toLowerCase();
            const w = clean(item.whatsapp).replace(/[+\s]/g, '');

            if (e && w) {
                // Sobrescribir: asumiendo que cualquier par válido es útil.
                // Si hay conflicto (mismo email, diferente whatsapp), gana el último procesado.
                // Podríamos priorizar 'contacts' si procesaramos list2 primero, pero da igual.
                uniqueMap.set(e, w);
            }
        });

        // Convertir a array
        const result = [];
        for (const [email, whatsapp] of uniqueMap) {
            result.push({ email, whatsapp });
        }

        return result;

    } catch (err) {
        console.error('Exception in getSupabaseContacts:', err);
        return [];
    }
}

module.exports = {
    getSupabaseContacts
};
