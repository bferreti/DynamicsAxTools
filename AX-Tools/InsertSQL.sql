DECLARE @PerfmonXML xml
SELECT @PerfmonXML = BulkColumn
FROM  OPENROWSET(BULK 'C:\Users\Administrator\Desktop\PerfmonTemplates\AxPerfmon_AOS.xml', SINGLE_BLOB) AS Template;
INSERT INTO AXTools_PerfmonTemplates ([SERVERTYPE],[ACTIVE],[TEMPLATEXML])
VALUES ('AOS', '1', @PerfmonXML)
