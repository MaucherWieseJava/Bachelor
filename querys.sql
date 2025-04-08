-- 1. Neue Spalten anlegen
ALTER TABLE [Name der Tabelle] ADD [Call Center] VARCHAR(10);
ALTER TABLE [Name der Tabelle] ADD [Kampagne] VARCHAR(50);
ALTER TABLE [Name der Tabelle] ADD [Meldedatum] DATE;
ALTER TABLE [Name der Tabelle] ADD [RKMDAT] INT;
ALTER TABLE [Name der Tabelle] ADD [DELLAT] INT;

-- 2. Call Center aus VP Vertragsnummer extrahieren (Zeichen 7 & 8)
UPDATE [Name der Tabelle]
SET [Call Center] = SUBSTRING([VP Vertragsnummer], 7, 2);

-- 3. Kampagne aufbauen aus Call Center + '-' + Campaign No_
UPDATE [Name der Tabelle]
SET [Kampagne] = [Call Center] + '-' + [Campaign No_];

-- 4. Meldedatum = Start_Insurance - 1 Tag
UPDATE [Name der Tabelle]
SET [Meldedatum] = DATEADD(DAY, -1, [Start_Insurance]);

-- 5. RKMDAT = 100 * Jahr + Monat von Meldedatum
UPDATE [Name der Tabelle]
SET [RKMDAT] = 100 * YEAR([Meldedatum]) + MONTH([Meldedatum]);

-- 6. DELLAT = 100 * Jahr + Monat von 'Deletion allowed at' + 1
UPDATE [Name der Tabelle]
SET [DELLAT] = 100 * YEAR([Deletion allowed at]) + MONTH([Deletion allowed at]) + 1;

-- 7. Finale SELECT-Abfrage für Export (z. B. in Excel)
SELECT 
    [VP Vertragsnummer],
    [Call Center],
    [Campaign No_],
    [Kampagne],
    [Start_Insurance],
    [Meldedatum],
    [RKMDAT],
    [Deletion allowed at],
    [DELLAT]
FROM [Name der Tabelle];