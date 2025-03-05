
ALTER TABLE deine_tabelle ADD Meldedatum DATE;

UPDATE deine_tabelle
SET Meldedatum = DATEADD(DAY, -1, Start_Insurance);


ALTER TABLE deine_tabelle ADD RKMDAT INT;

UPDATE deine_tabelle
SET RKMDAT = 100 * YEAR(Meldedatum) + MONTH(Meldedatum);