Lo script geolocalizza in automatico un elenco di indirizzi contenuto in un file Excel.

Funzionamento:

Apre File.xlsx (nella cartella corrente) e legge il primo foglio.
Per ogni riga (dalla 2 in poi), estrae indirizzo, CAP e località dalle colonne 5, 6 e 7.
Compone l'indirizzo completo (indirizzo località CAP ITALIA) e lo invia al servizio di geocoding Photon (tramite geopy) per ottenere latitudine e longitudine.
Se la geocodifica dell'indirizzo completo fallisce, effettua un secondo tentativo usando solo CAP + località come fallback.
Scrive il risultato nelle colonne 23, 24 e 25:
lat/lon trovate, precisione vuota → geocodifica riuscita con indirizzo completo
lat/lon trovate, precisione 0 → geocodifica riuscita solo col fallback (meno precisa)
lat/lon vuote, precisione -1 → geocodifica fallita in entrambi i tentativi
Rispetta un rate limit di almeno 1 secondo tra le richieste, per non sovraccaricare il servizio Photon pubblico.
Salva progressivamente il file ogni 10 righe elaborate (oltre che al termine), così da non perdere il lavoro in caso di interruzione su file di grandi dimensioni.
In sintesi: automatizza l'arricchimento di un elenco anagrafico/indirizzi con le coordinate geografiche, utile ad esempio per mappare clienti, sedi o punti vendita.
