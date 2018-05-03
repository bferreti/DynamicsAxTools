USE [DynamicsAxTools]
GO

INSERT INTO [dbo].[AXTools_EmailProfile]
           ([ID]
           ,[USERID]
           ,[SMTPSERVER]
           ,[SMTPPORT]
           ,[SMTPSSL]
           ,[FROM]
           ,[TO]
           ,[CC]
           ,[BCC])
     VALUES
           ('ID'
           ,'UserID' --ID from AXTools_UserAccount
           ,'SMTPSERVER'
           ,'SMTPPORT'
           ,'SMTPSSL' -- Uses SSL 1 or 0
           ,'FROM' --Address sending the email (with description)
           ,'TO' --Separate by ;
           ,'CC' --Separate by ;
           ,'BCC' --Separate by ;
		   )
GO


INSERT INTO [dbo].[AXTools_EmailProfile]
           ([ID]
           ,[USERID]
           ,[SMTPSERVER]
           ,[SMTPPORT]
           ,[SMTPSSL]
           ,[FROM]
           ,[TO]
           ,[CC]
           ,[BCC])
     VALUES
           ('GMAIL'
           ,'PS.MONITORING.PFE' --ID from AXTools_UserAccount
           ,'smtp.gmail.com'
           ,'587'
           ,1 -- Uses SSL 1 or 0
           ,'SQL Monitor <ps.monitoring.pfe@gmail.com>' --Address sending the email (with description)
           ,'bferreti@gmail.com;bferreti@microsoft.com' --Separate by ;
           ,'' --Separate by ;
           ,'' --Separate by ;
		   )
GO