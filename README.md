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

-------------------------------------------------------------------------------------------------------------------------------------------------------------

Photon (komoot) espone un endpoint REST pubblico su https://photon.komoot.io/api/, che è esattamente quello usato di default dalla classe Photon() di geopy nello script.

Esempio di chiamata equivalente a quella fatta dallo script (query = indirizzo completo):


curl -G "https://photon.komoot.io/api/" \
  --data-urlencode "q=Via Roma 1 Milano 20100 ITALIA" \
  --data-urlencode "limit=1"
Esempio con solo CAP + località (il fallback usato in caso di errore):


curl -G "https://photon.komoot.io/api/" \
  --data-urlencode "q=20100 Milano ITALIA" \
  --data-urlencode "limit=1"
Risposta (GeoJSON), da cui geopy estrae latitude/longitude:


{
  "features": [
    {
      "geometry": {
        "coordinates": [9.1900, 45.4642],
        "type": "Point"
      },
      "properties": {
        "name": "Via Roma",
        "city": "Milano",
        "postcode": "20100",
        "country": "Italy",
        ...
      },
      "type": "Feature"
    }
  ],
  "type": "FeatureCollection"
}
Nota: coordinates in GeoJSON è [longitude, latitude] — ordine invertito rispetto a location.latitude/location.longitude di geopy, che gestisce già la conversione correttamente.

Parametri utili aggiuntivi supportati dall'API: lang (es. it), lat/lon (per dare priorità geografica ai risultati), bbox (limitare a un'area).
