USE [master]
GO
CREATE DATABASE [DynamicsAxTools]
GO

ALTER DATABASE [DynamicsAxTools] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET ARITHABORT OFF 
GO

GO

ALTER DATABASE [DynamicsAxTools] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [DynamicsAxTools] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [DynamicsAxTools] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET  DISABLE_BROKER 
GO

ALTER DATABASE [DynamicsAxTools] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [DynamicsAxTools] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET HONOR_BROKER_PRIORITY OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET RECOVERY SIMPLE 
GO

ALTER DATABASE [DynamicsAxTools] SET  MULTI_USER 
GO

ALTER DATABASE [DynamicsAxTools] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [DynamicsAxTools] SET DB_CHAINING OFF 
GO

ALTER DATABASE [DynamicsAxTools] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO

ALTER DATABASE [DynamicsAxTools] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO

ALTER DATABASE [DynamicsAxTools] SET DELAYED_DURABILITY = DISABLED 
GO

USE [DynamicsAxTools]
GO

ALTER DATABASE [DynamicsAxTools] SET  READ_WRITE 
GO

USE [DynamicsAxTools]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXInstall_Status](
       [GUID] [nvarchar](36) NULL,
       [SERVERNAME] [nvarchar](50) NULL,
       [TYPE] [nvarchar](25) NULL,
       [STATUS] [nvarchar](25) NULL,
       [LOG] [nvarchar](254) NULL,
       [MODIFIEDDATETIME] [datetime] NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [IX_CREATEDDATETIME_GUID] ON [dbo].[AXInstall_Status]
(
       [CREATEDDATETIME] ASC,
       [GUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXInstall_StatusDetails](
       [GUID] [nvarchar](36) NULL,
       [SERVERNAME] [nvarchar](50) NULL,
       [PACKAGENAME] [nvarchar](150) NULL,
       [TYPE] [nvarchar](2) NULL,
       [STATUS] [nvarchar](25) NULL,
       [ARGUMENTLIST] [nvarchar](max) NULL,
       [LOG] [nvarchar](max) NULL,
       [LOGDATE] [nvarchar](25) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [IX_CREATEDDATETIME_GUID] ON [dbo].[AXInstall_StatusDetails]
(
       [CREATEDDATETIME] ASC,
       [GUID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_AXBatchJobs](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [STARTDATETIME] [datetime] NULL,
       [ENDDATETIME] [datetime] NULL,
       [CAPTION] [nvarchar](200) NULL,
       [STATUS] [nvarchar](15) NULL,
       [CREATEDBY] [nvarchar](15) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_AXNumberSequences](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [NUMBERSEQUENCE] [nvarchar](20) NULL,
       [TXT] [nvarchar](120) NULL,
       [FORMAT] [nvarchar](40) NULL,
       [STATUS] [nvarchar](15) NULL,
       [CONTINUOUS] [tinyint] NULL,
       [TRANSID] [bigint] NULL,
       [SESSIONID] [int] NULL,
       [USERID] [nvarchar](16) NULL,
       [MODIFIEDBY] [nvarchar](16) NULL,
       [SESSIONLOGINDATETIME] [datetime] NULL,
       [MODIFIEDDATETIME] [datetime] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_ExecutionLog](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [CPU] [decimal](6, 4) NULL,
       [BLOCKING] [int] NULL,
       [WAITING] [int] NULL,
       [GRD] [tinyint] NULL,
       [GRDTOTAL] [int] NULL,
       [STATS] [tinyint] NULL,
       [STATSTOTAL] [int] NULL,
       [EMAIL] [tinyint] NULL,
       [REPORT] [nvarchar](200) NULL,
       [GUID] [nvarchar](36) NULL,
       [LOG] [nvarchar](500) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_GRDLog](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [TABLENAME] [nvarchar](50) NULL,
       [STATSTYPE] [nvarchar](8) NULL,
       [STATEMENT] [nvarchar](150) NULL,
       [JOBNAME] [nvarchar](150) NULL,
       [STARTED] [datetime] NULL,
       [FINISHED] [datetime] NULL,
       [LOG] [nvarchar](500) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_GRDStatistics](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [TABLENAME] [sysname] NOT NULL,
       [INDEXNAME] [sysname] NOT NULL,
       [INDEXID] [int] NOT NULL,
       [ROWSTOTAL] [bigint] NULL,
       [ROWSMODIFIED] [bigint] NULL,
       [SIZEMB] [decimal](12, 2) NULL,
       [PERCENTCHANGE] [bigint] NULL,
       [LASTUPDATE] [datetime] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [idx_CreatedDateTime_Environment] ON [dbo].[AXMonitor_GRDStatistics]
(
       [CREATEDDATETIME] ASC,
       [ENVIRONMENT] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_PerfmonData](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [PATH] [nvarchar](300) NULL,
       [VALUE] [decimal](16, 2) NULL,
       [TIMESTAMP] [datetime] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_SQLConfiguration](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [DISPLAYNAME] [nvarchar](50) NULL,
       [DESCRIPTION] [nvarchar](150) NULL,
       [RUNVALUE] [nvarchar](10) NULL,
       [CONFIGVALUE] [nvarchar](10) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_SQLInformation](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [NAME] [nvarchar](50) NULL,
       [VALUE] [nvarchar](300) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXMonitor_SQLRunningSpids](
       [ENVIRONMENT] [nvarchar](30) NULL,
       [START_DATE_TIME] [datetime] NULL,
       [SPID] [nvarchar](10) NULL,
       [BLOCKER] [nvarchar](10) NULL,
       [STATUS] [nvarchar](10) NULL,
       [HOST_NAME] [nvarchar](50) NULL,
       [CONTEXT_INFO] [nvarchar](50) NULL,
       [WAIT_TIME_MS] [bigint] NULL,
       [TOTAL_TIME_MS] [bigint] NULL,
       [CPU_TIME_MS] [bigint] NULL,
       [CPU_TIME_PERC] [decimal](15, 10) NULL,
       [READS] [bigint] NULL,
       [WRITES] [bigint] NULL,
       [LOGICAL_READS] [bigint] NULL,
       [WAIT_TYPE] [nvarchar](50) NULL,
       [DATABASE] [nvarchar](50) NULL,
       [SQL_TEXT] [nvarchar](max) NULL,
       [PLAN_HANDLE] [nvarchar](max) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate())
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [idx_CreatedDateTime_Environment] ON [dbo].[AXMonitor_SQLRunningSpids]
(
       [CREATEDDATETIME] ASC,
       [ENVIRONMENT] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_AOSServices](
       [SERVERNAME] [nvarchar](30) NULL,
       [SERVERTYPE] [nvarchar](10) NULL,
       [MACHINENAME] [nvarchar](50) NULL,
       [NAME] [nvarchar](50) NULL,
       [STATUS] [nvarchar](30) NULL,
       [DISPLAYNAME] [nvarchar](150) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [Clustered_ReportID_ServerName_ServerType] ON [dbo].[AXReport_AOSServices]
(
       [REPORTDATE] ASC,
       [SERVERNAME] ASC,
       [SERVERTYPE] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_BatchJobs](
       [HISTORYCAPTION] [nvarchar](150) NULL,
       [JOBCAPTION] [nvarchar](150) NULL,
       [STATUS] [nvarchar](30) NULL,
       [SERVERID] [nvarchar](30) NULL,
       [STARTDATETIMECST] [datetime] NULL,
       [ENDDATETIMECST] [datetime] NULL,
       [EXECUTEDBY] [nvarchar](15) NULL,
       [BATCHID] [bigint] NULL,
       [BATCHJOBID] [bigint] NULL,
       [BATCHJOBHISTORYID] [bigint] NULL,
       [REPORTDATE] [date] NULL,
       [LOG] [nvarchar](max) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_CDXJobs](
       [JOBID] [nvarchar](10) NULL,
       [DATASTORESTATUS] [bigint] NULL,
       [STATUSDOWNLOADSESSIONDATASTORE] [nvarchar](30) NULL,
       [MESSAGE] [nvarchar](max) NULL,
       [DATEREQUESTED] [datetime] NULL,
       [DATEDOWNLOADED] [datetime] NULL,
       [DATEAPPLIED] [datetime] NULL,
       [CURRENTROWVERSION] [bigint] NULL,
       [ROWSAFFECTED] [bigint] NULL,
       [DATAFILEOUTPUTPATH] [nvarchar](max) NULL,
       [SESSIONSTATUS] [bigint] NULL,
       [STATUSDOWNLOADSESSION] [nvarchar](30) NULL,
       [DATABASE_] [nvarchar](30) NULL,
       [NAME] [nvarchar](50) NULL,
       [MODIFIEDDATETIME] [datetime] NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_DBInstances](
       [SERVERNAME] [nvarchar](30) NULL,
       [SERVERTYPE] [nvarchar](10) NULL,
       [DBSERVER] [nvarchar](50) NULL,
       [DBNAME] [nvarchar](50) NULL,
       [DETAILS] [nvarchar](255) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [Clustered_ReportID_ServerName_ServerType] ON [dbo].[AXReport_DBInstances]
(
       [REPORTDATE] ASC,
       [SERVERNAME] ASC,
       [SERVERTYPE] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_EventLogs](
       [SERVERNAME] [nvarchar](30) NULL,
       [SERVERTYPE] [nvarchar](10) NULL,
       [MACHINENAME] [nvarchar](50) NULL,
       [LOGNAME] [nvarchar](50) NULL,
       [ENTRYTYPE] [nvarchar](50) NULL,
       [EVENTID] [bigint] NULL,
       [SOURCE] [nvarchar](max) NULL,
       [TIMEGENERATED] [datetime] NULL,
       [MESSAGE] [nvarchar](max) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [Clustered_ReportID_ServerName_ServerType] ON [dbo].[AXReport_EventLogs]
(
       [REPORTDATE] ASC,
       [SERVERNAME] ASC,
       [SERVERTYPE] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_LongBatchJobs](
       [JOB] [nvarchar](150) NULL,
       [COUNT] [int] NULL,
       [STATUS] [nvarchar](30) NULL,
       [DURATION] [int] NULL,
       [EXECUTEDBY] [nvarchar](15) NULL,
       [SERVERID] [nvarchar](30) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_MRP](
       [REQPLANID] [nvarchar](30) NULL,
       [STARTDATETIME] [datetime] NULL,
       [ENDDATETIME] [datetime] NULL,
       [CANCELLED] [int] NULL,
       [USEDCHILDTHREADS] [int] NULL,
       [MAXCHILDTHREADS] [int] NULL,
       [COMPLETEUPDATE] [int] NULL,
       [USEDTODAYSDATE] [datetime] NULL,
       [NUMOFITEMS] [bigint] NULL,
       [NUMOFINVENTONHAND] [bigint] NULL,
       [NUMOFSALESLINE] [bigint] NULL,
       [NUMOFPURCHLINE] [bigint] NULL,
       [NUMOFTRANSFERPLANNEDORDER] [bigint] NULL,
       [NUMOFITEMPLANNEDORDER] [bigint] NULL,
       [NUMOFINVENTJOURNAL] [bigint] NULL,
       [TIMECOPY] [bigint] NULL,
       [TIMECOVERAGE] [bigint] NULL,
       [TIMEUPDATE] [bigint] NULL,
       [REPORTDATE] [date] NULL,
       [LOG] [nvarchar](max) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_NetwokConn](
       [ENVIRONMENT] [nvarchar](50) NULL,
       [SERVERNAME] [nvarchar](30) NULL,
       [IP] [nvarchar](50) NULL,
       [PINGAVG] [float] NULL,
       [LOSTPACKAGE] [int] NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_Netwoking](
       [ENVIRONMENT] [nvarchar](50) NULL,
       [SOURCENAME] [nvarchar](30) NULL,
       [SERVERNAME] [nvarchar](30) NULL,
       [IP] [nvarchar](50) NULL,
       [AVERAGE] [float] NULL,
       [MAXIMUM] [float] NULL,
       [MINIMUM] [float] NULL,
       [SENT] [int] NULL,
       [LOST] [int] NULL,
       [LOSS] [float] NULL,
       [MESSAGE] [nvarchar](254) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_PerfmonData](
       [SERVERNAME] [nvarchar](30) NULL,
       [SERVERTYPE] [nvarchar](10) NULL,
       [COUNTERTYPE] [nvarchar](10) NULL,
       [REPORTVIEW] [bit] NULL,
       [PATH] [nvarchar](max) NULL,
       [MAXIMUM] [float] NULL,
       [MINIMUM] [float] NULL,
       [AVERAGE] [float] NULL,
       [FULLPATH] [nvarchar](max) NULL,
       [STARTDATETIME] [datetime] NULL,
       [ENDDATETIME] [datetime] NULL,
       [SAMPLES] [bigint] NULL,
       [COUNTER] [nvarchar](max) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_ProcessDetail](
       [SERVERNAME] [nvarchar](30) NULL,
       [SERVERTYPE] [nvarchar](10) NULL,
       [MACHINENAME] [nvarchar](50) NULL,
       [NAME] [nvarchar](50) NULL,
       [ID] [bigint] NULL,
       [HANDLES] [bigint] NULL,
       [VM] [bigint] NULL,
       [WS] [bigint] NULL,
       [PM] [bigint] NULL,
       [NPM] [bigint] NULL,
       [WORKINGSET] [bigint] NULL,
       [PAGEDMEMORYSIZE] [bigint] NULL,
       [PRIVATEMEMORYSIZE] [bigint] NULL,
       [VIRTUALMEMORYSIZE] [bigint] NULL,
       [BASEPRIORITY] [bigint] NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_PADDING ON

GO
CREATE CLUSTERED INDEX [Clustered_ReportID_ServerName_ServerType] ON [dbo].[AXReport_ProcessDetail]
(
       [REPORTDATE] ASC,
       [SERVERNAME] ASC,
       [SERVERTYPE] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_SQLServerLogs](
       [LOGDATE] [datetime] NULL,
       [PROCESSINFO] [nvarchar](50) NULL,
       [TEXT] [nvarchar](max) NULL,
       [SERVER] [nvarchar](50) NULL,
       [DATABASE] [nvarchar](50) NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXReport_SSRSLogs](
       [STATUS] [nvarchar](50) NULL,
       [INSTANCENAME] [nvarchar](30) NULL,
       [REPORTPATH] [nvarchar](max) NULL,
       [USERNAME] [nvarchar](30) NULL,
       [FORMAT] [nvarchar](30) NULL,
       [TIMESTART] [datetime] NULL,
       [TIMEEND] [datetime] NULL,
       [TIMEDATARETRIEVAL] [bigint] NULL,
       [TIMEPROCESSING] [bigint] NULL,
       [TIMERENDERING] [bigint] NULL,
       [REPORTDATE] [date] NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_AccountProfile](
       [ID] [nvarchar](60) NOT NULL,
       [DATA] [nvarchar](max) NULL,
       [CREATEDDATETIME] [datetime] NULL,
CONSTRAINT [PK_AXTools_AccountProfile] PRIMARY KEY CLUSTERED 
(
       [ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_EmailLog](
       [SENT] [tinyint] NOT NULL,
       [EMAILPROFILE] [nvarchar](60) NOT NULL,
       [SUBJECT] [nvarchar](200) NULL,
       [BODY] [nvarchar](max) NULL,
       [ATTACHMENT] [nvarchar](200) NULL,
       [LOG] [nvarchar](500) NULL,
       [GUID] [nvarchar](36) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_EmailProfile](
       [PROFILEID] [nvarchar](60) NOT NULL,
       [CONNECTIONID] [nvarchar](60) NULL,
       [CONNECTIONINFO] [nvarchar](max) NULL,
       [FROM] [nvarchar](max) NULL,
       [TO] [nvarchar](max) NULL,
       [CC] [nvarchar](max) NULL,
       [BCC] [nvarchar](max) NULL,
       [CREATEDDATETIME] [datetime] NULL,
CONSTRAINT [PK_AXTools_EmailProfile] PRIMARY KEY CLUSTERED 
(
       [PROFILEID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_Environments](
       [ENVIRONMENT] [nvarchar](30) NOT NULL,
       [DESCRIPTION] [nvarchar](100) NULL,
       [DBSERVER] [nvarchar](50) NULL,
       [DBNAME] [nvarchar](50) NULL,
       [CPUTHOLD] [int] NULL,
       [BLOCKTHOLD] [int] NULL,
       [WAITINGTHOLD] [int] NULL,
       [EMAIL] [tinyint] NULL,
       [EMAILPROFILE] [nvarchar](60) NULL,
       [GRD] [tinyint] NULL,
       [STATS] [tinyint] NULL,
       [DYNPERFSERVER] [nvarchar](50) NULL,
       [DYNPERFNAME] [nvarchar](50) NULL,
       [CREATEDDATETIME] [datetime] NULL DEFAULT (getdate()),
CONSTRAINT [PK_AXTools_Environments] PRIMARY KEY CLUSTERED 
(
       [ENVIRONMENT] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_ExecutionLog](
       [CREATEDDATETIME] [datetime] NULL,
       [GUID] [nvarchar](36) NULL,
       [LOG] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
CREATE CLUSTERED INDEX [idx_Clustered_CreatedDateTime] ON [dbo].[AXTools_ExecutionLog]
(
       [CREATEDDATETIME] DESC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, DATA_COMPRESSION = PAGE) ON [PRIMARY]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_PerfmonTemplates](
       [SERVERTYPE] [nvarchar](50) NULL,
       [ACTIVE] [bit] NULL,
       [TEMPLATEXML] [xml] NULL,
       [TEMPLATETXT] [nvarchar](max) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
WITH
(
DATA_COMPRESSION = PAGE
)

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AXTools_Servers](
       [ENVIRONMENT] [nvarchar](30) NOT NULL,
       [ACTIVE] [tinyint] NULL,
       [SERVERNAME] [nvarchar](50) NOT NULL,
       [SERVERTYPE] [nvarchar](50) NOT NULL,
       [IP] [nvarchar](50) NULL,
       [DOMAIN] [nvarchar](50) NULL,
       [FQDN] [nvarchar](50) NULL,
       [CREATEDDATETIME] [datetime] NULL
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[AXInstall_Status] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXInstall_StatusDetails] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXMonitor_AXNumberSequences] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_AOSServices] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_BatchJobs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_CDXJobs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_DBInstances] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_EventLogs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_LongBatchJobs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_MRP] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_NetwokConn] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_Netwoking] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_PerfmonData] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_ProcessDetail] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_SQLServerLogs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXReport_SSRSLogs] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_AccountProfile] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_EmailLog] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_EmailProfile] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_ExecutionLog] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_PerfmonTemplates] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
ALTER TABLE [dbo].[AXTools_EmailProfile]  WITH CHECK ADD  CONSTRAINT [FK_AXTools_EmailProfile_AXTools_AccountProfile] FOREIGN KEY([CONNECTIONID])
REFERENCES [dbo].[AXTools_AccountProfile] ([ID])
GO
ALTER TABLE [dbo].[AXTools_EmailProfile] CHECK CONSTRAINT [FK_AXTools_EmailProfile_AXTools_AccountProfile]
GO
ALTER TABLE [dbo].[AXTools_Environments]  WITH NOCHECK ADD  CONSTRAINT [FK_AXTools_Environments_AXTools_EmailProfile] FOREIGN KEY([EMAILPROFILE])
REFERENCES [dbo].[AXTools_EmailProfile] ([PROFILEID])
GO
ALTER TABLE [dbo].[AXTools_Environments] CHECK CONSTRAINT [FK_AXTools_Environments_AXTools_EmailProfile]
GO
ALTER TABLE [dbo].[AXTools_Servers]  WITH NOCHECK ADD  CONSTRAINT [FK_AXTools_Servers_AXTools_Environments] FOREIGN KEY([ENVIRONMENT])
REFERENCES [dbo].[AXTools_Environments] ([ENVIRONMENT])
GO
ALTER TABLE [dbo].[AXTools_Servers] CHECK CONSTRAINT [FK_AXTools_Servers_AXTools_Environments]
GO
ALTER TABLE [dbo].[AXTools_Servers] ADD  DEFAULT (getdate()) FOR [CREATEDDATETIME]
GO
