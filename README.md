Dette skriptet behandler flere Excel-filer ved å utføre spesifikke datamanipulasjoner og lagrer resultatet i en ny Excel-fil. Skriptet er designet for å være brukervennlig og bruker GUI-dialoger for filvalg og lagringssted ved hjelp av Tkinter-biblioteket.

Forutsetninger
For å kjøre dette skriptet trenger du Python 3.x installert på systemet ditt sammen med følgende Python-pakker: pandas og tkinter.

Funksjonalitet
Skriptet utfører følgende operasjoner:

Ekstraher dato fra filnavn: En funksjon som ekstraherer dato fra filnavnet og konverterer den til formatet 'dd.mm.yyyy'.

Filvalg: Skriptet bruker en filvelger-dialog for å la brukeren velge flere Excel-filer (.xlsx).

Les og prosesser hver Excel-fil:

Laster inn Excel-filen og leser det første arket.
Finn raden der "Tips" er plassert i kolonne A og beregn målet for kolonnen (ti kolonner til høyre for kolonne A).
Finn rader der "25%" og "Betalingsformidling" er plassert i kolonne A og ekstraher relevante verdier.
Traverser kolonnen under og til høyre for "Betalingsformidling" for å beregne summer for spesifikke beskrivelse (f.eks. uintegrert kontanter, Vipps).
Finn raden der "Endring i kredittbalanse" er plassert i kolonne A og ekstraher verdien.
Transformering av data:

Lag en DataFrame fra de ekstraherte dataene.
Iterer gjennom radene og legg til relevante data i en ny DataFrame, med formatering av beløpene.
Lagre prosesserte data til ny Excel-fil:

Bruker en filvelger-dialog for å la brukeren velge lagringssted for den nye Excel-filen.
Lagre de transformerede dataene i en ny Excel-fil med ett ark kalt "Summary".
Suksessmelding: En melding skrives til konsollen når filen er lagret.

Bruk
Kjør Skriptet: Kjør skriptet i et Python-miljø.
Velg Inndatafiler: En filvelger-dialog vil dukke opp. Velg de Excel-filene (.xlsx) du ønsker å behandle.
Lagre Utdatapfilen: Etter behandling vil en annen filvelger-dialog dukke opp. Velg hvor du vil lagre den nye Excel-filen.
Fullføring: En melding i konsollen bekrefter at filen er lagret.
