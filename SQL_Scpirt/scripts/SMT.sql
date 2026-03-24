USE [master]
GO
/****** Object:  Database [OK2SHIP_SMT]    Script Date: 3/20/2026 5:31:44 PM ******/
CREATE DATABASE [OK2SHIP_SMT]
 use[OK2SHIP_SMT]
CREATE TABLE [dbo].[ACCOUNT](
	[User_ID] [int] NOT NULL,
	[Username] [nvarchar](50) NOT NULL,
	[Password] [nvarchar](50) NOT NULL,
	[IsActive] [int] NOT NULL,
	[Role] [int] NOT NULL,
	[CreateBy] [int] NULL,
	[CreateAt] [datetime] NULL,
	[ModifyDate] [datetime] NULL,
	[ModifyAt] [int] NULL,
	[Information] [nvarchar](max) NULL,
 CONSTRAINT [PK_ACCOUNT] PRIMARY KEY CLUSTERED 
(
	[User_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_BONDING]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_BONDING](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Image_Before] [varbinary](max) NULL,
	[Image_After] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_BONDING_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_BONDING_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Remark] [nvarchar](30) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_ACF_BONDING_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_FLATNESS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_FLATNESS](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Point1] [nvarchar](max) NULL,
	[Point2] [nvarchar](max) NULL,
	[Point3] [nvarchar](max) NULL,
	[Point4] [nvarchar](max) NULL,
	[Point5] [nvarchar](max) NULL,
	[Point6] [nvarchar](max) NULL,
	[Point7] [nvarchar](max) NULL,
	[Point8] [nvarchar](max) NULL,
	[Point9] [nvarchar](max) NULL,
	[Point10] [nvarchar](max) NULL,
	[Flatness] [nvarchar](max) NULL,
	[Judgement] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_FLATNESS_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_FLATNESS_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_ACF_FLATNESS_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_WCA]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_WCA](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Data] [nvarchar](max) NULL,
	[Type] [nvarchar](50) NULL,
	[Operator] [nchar](10) NULL,
 CONSTRAINT [PK_ACF_WCA] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ACF_WETTING]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ACF_WETTING](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Machine] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Side] [nvarchar](max) NULL,
	[After_Plasma] [nvarchar](max) NULL,
	[Before_Packing] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_AVAILABLE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_AVAILABLE](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Status] [int] NULL,
 CONSTRAINT [PK_AIR_BUBBLE_AVAILABLE] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_BEFORE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_BEFORE](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Picture] [varbinary](max) NULL,
	[Data] [float] NULL,
	[KeyDic] [nvarchar](50) NULL,
 CONSTRAINT [PK_AIR_BUBBLE_BEFORE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_BEFORE_PIC]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_BEFORE_PIC](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Image Before$CONVERTER] [varbinary](max) NULL,
	[Image Before+$CONVERTER] [varbinary](max) NULL,
	[ProductID] [nvarchar](50) NULL,
 CONSTRAINT [PK_AIR_BUBBLE_BEFORE_PIC] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_COMMENT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_COMMENT](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[TapeName] [nvarchar](50) NOT NULL,
	[Comment1] [nvarchar](max) NULL,
	[Comment2] [nvarchar](max) NULL,
 CONSTRAINT [PK_AIR_BUBBLE_COMMENT] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_DATA]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_DATA](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[TONG] [varbinary](max) NULL,
	[Measure TONG (mm2)] [float] NULL,
	[% Area TONG] [nvarchar](50) NULL,
	[TRU] [varbinary](max) NULL,
	[Measure TRU (mm2)] [float] NULL,
	[% Area TRU] [nvarchar](50) NULL,
	[Judgement] [nchar](10) NULL,
	[KeyDIC] [nvarchar](50) NULL,
	[Measure(mm)] [nvarchar](20) NULL,
	[Adhesive] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AIR_BUBBLE_REFER]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AIR_BUBBLE_REFER](
	[ID] [int] NOT NULL,
	[ItemMain] [nchar](10) NULL,
	[LotMain] [nchar](10) NULL,
	[ItemCodeRefer] [nchar](10) NULL,
	[LotNoRefer] [nchar](10) NULL,
 CONSTRAINT [PK_AIR_BUBBLE_REFER] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ASSY_YIELD]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ASSY_YIELD](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Area] [uniqueidentifier] NULL,
	[Process] [nvarchar](max) NULL,
	[Top_SMT] [nvarchar](max) NULL,
	[Top_Backend] [nvarchar](max) NULL,
	[OQC] [nvarchar](max) NULL,
 CONSTRAINT [PK_ASSY_YIELD] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ASSY_YIELD_IMAGE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ASSY_YIELD_IMAGE](
	[ID] [int] NOT NULL,
	[Image] [varbinary](max) NULL,
	[Area] [uniqueidentifier] NULL,
 CONSTRAINT [PK_ASSY_YIELD_IMAGE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BAR_CODE_VERIFICATION]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BAR_CODE_VERIFICATION](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[ListSN] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CAMERA_SMT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CAMERA_SMT](
	[ID] [int] NOT NULL,
	[Process] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CROSS_SECTION]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CROSS_SECTION](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image1] [varbinary](max) NULL,
	[Image2] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CROSS_SECTION_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CROSS_SECTION_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image1] [varbinary](max) NULL,
	[Image2] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CROSS_SECTION_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CROSS_SECTION_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_CROSS_SECTION_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[EDIT_HISTORY]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EDIT_HISTORY](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data_Before] [nvarchar](max) NULL,
	[Data_After] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[Date_Modify] [nvarchar](max) NULL,
	[Process] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ENVIRONMENT_EN_DURANCE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ENVIRONMENT_EN_DURANCE](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[Process] [nvarchar](30) NULL,
 CONSTRAINT [PK_ENVIRONMENT_EN_DURANCE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ENVIRONMENT_EN_DURANCE_IMG]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ENVIRONMENT_EN_DURANCE_IMG](
	[ID] [int] NOT NULL,
	[Area] [int] NOT NULL,
	[Image] [varbinary](max) NULL,
 CONSTRAINT [PK_ENVIRONMENT_EN_DURANCE_IMG] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ENVIRONMENT_EN_DURANCE_IMG_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ENVIRONMENT_EN_DURANCE_IMG_NAS](
	[ID] [int] NOT NULL,
	[Area] [nchar](10) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](max) NULL,
 CONSTRAINT [PK_ENVIRONMENT_EN_DURANCE_IMG_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FAI_Auto]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FAI_Auto](
	[ItemCode] [nvarchar](50) NULL,
	[LotNo] [nvarchar](50) NULL,
	[Operator] [nvarchar](50) NULL,
	[Machine] [nvarchar](50) NULL,
	[MDate] [nvarchar](50) NULL,
	[FAI_No] [nvarchar](50) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL,
	[ID] [nvarchar](max) NULL,
	[FAI_Data] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FAI_Spec]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FAI_Spec](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[NormDim] [nvarchar](max) NULL,
	[TolMax] [nvarchar](max) NULL,
	[TolMin] [nvarchar](max) NULL,
	[Instrument] [nvarchar](max) NULL,
	[FAI_No] [nvarchar](max) NULL,
	[Distribution] [nvarchar](max) NULL,
	[SheetNo] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FLEX_BENDING]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FLEX_BENDING](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Sel_report] [nvarchar](max) NULL,
	[Before] [nvarchar](max) NULL,
	[After_1] [nvarchar](max) NULL,
	[After_2] [nvarchar](max) NULL,
	[After_3] [nvarchar](max) NULL,
	[After_4] [nvarchar](max) NULL,
	[After_5] [nvarchar](max) NULL,
	[After_10] [nvarchar](max) NULL,
	[After_15] [nvarchar](max) NULL,
	[After_20] [nvarchar](max) NULL,
	[After_25] [nvarchar](max) NULL,
	[After_30] [nvarchar](max) NULL,
	[After_40] [nvarchar](max) NULL,
	[After_50] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FLEX_BENDING_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FLEX_BENDING_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Cycles_name] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[ItemName] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FLEX_BENDING_LOGFILE_OLD]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FLEX_BENDING_LOGFILE_OLD](
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Cycles_name] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[ItemName] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FLEX_BENDING_LOGFILE_S2]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FLEX_BENDING_LOGFILE_S2](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](10) NULL,
	[LotNo] [nvarchar](10) NULL,
	[Net_No] [nvarchar](50) NULL,
	[Pcs_No] [nchar](20) NULL,
	[Data] [nvarchar](50) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Cycles_name] [nvarchar](10) NULL,
	[Remark] [nvarchar](10) NULL,
	[Shift] [nvarchar](10) NULL,
	[UserID] [nvarchar](10) NULL,
	[ItemName] [nvarchar](10) NULL,
 CONSTRAINT [PK_FLEX_BENDING_LOGFILE_S2] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FLEX_BENDING_S2]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FLEX_BENDING_S2](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Net_No] [nchar](10) NULL,
	[Pcs_No] [nchar](20) NULL,
	[Sel_report] [nchar](10) NULL,
	[Before] [nchar](10) NULL,
	[After_1] [nchar](10) NULL,
	[After_2] [nchar](10) NULL,
	[After_3] [nchar](10) NULL,
	[After_4] [nchar](10) NULL,
	[After_5] [nchar](10) NULL,
	[After_10] [nchar](10) NULL,
	[After_15] [nchar](10) NULL,
	[After_20] [nchar](10) NULL,
	[After_25] [nchar](10) NULL,
	[After_30] [nchar](10) NULL,
	[After_40] [nchar](10) NULL,
	[After_50] [nchar](10) NULL,
	[Logfile] [nchar](10) NULL,
	[Remark] [nchar](10) NULL,
	[Shift] [nchar](10) NULL,
 CONSTRAINT [PK_FLEX_BENDING_S2] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GAP_CONNECTOR]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GAP_CONNECTOR](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Image1] [varbinary](max) NULL,
	[Image2] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GAP_CONNECTOR_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GAP_CONNECTOR_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Image1] [varbinary](max) NULL,
	[Image2] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GAP_CONNECTOR_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GAP_CONNECTOR_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_GAP_CONNECTOR_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[HEAT_SOAK_AND_BEND]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[HEAT_SOAK_AND_BEND](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Sel_report] [nvarchar](max) NULL,
	[Before] [nvarchar](max) NULL,
	[After_1] [nvarchar](max) NULL,
	[After_2] [nvarchar](max) NULL,
	[After_3] [nvarchar](max) NULL,
	[After_4] [nvarchar](max) NULL,
	[After_5] [nvarchar](max) NULL,
	[After_10] [nvarchar](max) NULL,
	[After_15] [nvarchar](max) NULL,
	[After_20] [nvarchar](max) NULL,
	[After_25] [nvarchar](max) NULL,
	[After_30] [nvarchar](max) NULL,
	[After_40] [nvarchar](max) NULL,
	[After_50] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[HEAT_SOAK_AND_BEND_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[HEAT_SOAK_AND_BEND_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Cycles_name] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[ItemName] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMPEDANCE_GRAPH]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMPEDANCE_GRAPH](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Image_Graph] [varbinary](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Zone] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMPEDANCE_IMAGE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMPEDANCE_IMAGE](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Image_Data] [nchar](10) NULL,
	[Image_Graph] [nchar](10) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMPEDANCE_SPEC]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMPEDANCE_SPEC](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Range] [nvarchar](max) NULL,
	[USL] [nvarchar](max) NULL,
	[LSL] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMPEDANCE_VAL]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMPEDANCE_VAL](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Zone] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IMPEDANCE_VAL_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IMPEDANCE_VAL_LOGFILE](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Zone] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_LINER_PEELING_COUPON]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_LINER_PEELING_COUPON](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_LINER_PEELING_COUPON_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_LINER_PEELING_COUPON_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_LINER_PEELING_COUPON_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_LINER_PEELING_COUPON_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](30) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_PSA_PEELING_COUPON]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_PSA_PEELING_COUPON](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_PSA_PEELING_COUPON_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_PSA_PEELING_COUPON_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_PSA_PEELING_COUPON_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_PSA_PEELING_COUPON_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](30) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_UNMATING_PULL_TEST]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_UNMATING_PULL_TEST](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_UNMATING_PULL_TEST_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_UNMATING_PULL_TEST_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IQC_UNMATING_PULL_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IQC_UNMATING_PULL_TEST_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](30) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_IQC_UNMATING_PULL_TEST_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Items_Details]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Items_Details](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[Items] [nvarchar](max) NULL,
	[Address] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[LINER_PEEL_TEST_ON_PRODUCT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LINER_PEEL_TEST_ON_PRODUCT](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[LINER_PEEL_TEST_ON_PRODUCT_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LINER_PEEL_TEST_ON_PRODUCT_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[LINER_PEEL_TEST_ON_PRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LINER_PEEL_TEST_ON_PRODUCT_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_LINER_PEEL_TEST_ON_PRODUCT_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[MATING_PULL_TEST]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MATING_PULL_TEST](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[MATING_PULL_TEST_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MATING_PULL_TEST_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[MATING_PULL_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MATING_PULL_TEST_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_MATING_PULL_TEST_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Modify_Log]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Modify_Log](
	[ID] [int] NOT NULL,
	[User_ID] [nvarchar](max) NULL,
	[User_Action] [nvarchar](max) NULL,
	[Before_Val] [nvarchar](max) NULL,
	[After_Val] [nvarchar](max) NULL,
	[FAI_No] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[NET_SPEC]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NET_SPEC](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[Sel_Report] [nvarchar](max) NULL,
	[Net_Name] [nvarchar](max) NULL,
	[Point+V] [nvarchar](max) NULL,
	[Point-V] [nvarchar](max) NULL,
	[Net_Internal] [nvarchar](max) NULL,
	[USL] [nvarchar](max) NULL,
	[LSL] [nvarchar](max) NULL,
	[SEI_USL] [nvarchar](max) NULL,
	[SEI_LSL] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[NET_SPEC1]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NET_SPEC1](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[SeI_Report] [nvarchar](max) NULL,
	[Net_Name] [nvarchar](max) NULL,
	[Point+V] [nvarchar](max) NULL,
	[Point-V] [nvarchar](max) NULL,
	[Net_Internal] [nvarchar](max) NULL,
	[USL] [nvarchar](max) NULL,
	[LSL] [nvarchar](max) NULL,
	[SEI_USL] [nvarchar](max) NULL,
	[SEI_LSL] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[OQC_B2B_Mating_Unmating]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[OQC_B2B_Mating_Unmating](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[lotNo] [nchar](10) NOT NULL,
	[T0] [varbinary](max) NULL,
	[T1] [varbinary](max) NULL,
	[T30] [varbinary](max) NULL,
	[T0U] [varbinary](max) NULL,
	[T1U] [varbinary](max) NULL,
	[T30U] [varbinary](max) NULL,
	[Force] [float] NULL,
	[FailureMode] [nvarchar](50) NULL,
	[Judgement] [nvarchar](50) NULL,
	[Graph] [varbinary](max) NULL,
	[ProductID] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[OQC_B2B_Mating_Unmating_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[OQC_B2B_Mating_Unmating_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[LoactionImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_OQC_B2B_Mating_Unmating_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PACKAGING_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PACKAGING_LOGFILE](
	[ID] [int] NOT NULL,
	[ShippingTo] [nchar](10) NOT NULL,
	[TrayCode] [nchar](10) NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[FlexTop$Image] [nvarchar](50) NULL,
	[FlexBottom$Image] [nvarchar](50) NULL,
	[Tray$Image] [nvarchar](50) NULL,
	[TrayAL$Image] [nvarchar](50) NULL,
	[ALBag$Image] [nvarchar](50) NULL,
	[CartonBox$Image] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationIMG] [nvarchar](200) NULL,
 CONSTRAINT [PK_PACKAGING_LOGFILE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PEEL_TEST]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PEEL_TEST](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PEEL_TEST_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PEEL_TEST_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PEEL_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PEEL_TEST_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_PEEL_TEST_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PEEL_TEST_WITHOUT_SUS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PEEL_TEST_WITHOUT_SUS](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PEEL_TEST_WITHOUT_SUS_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PEEL_TEST_WITHOUT_SUS_NAS](
	[ID] [nchar](10) NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nchar](10) NULL,
	[Data] [nchar](10) NULL,
	[LocationImg] [nchar](10) NULL,
 CONSTRAINT [PK_PEEL_TEST_WITHOUT_SUS_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PRODUCT_ID]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PRODUCT_ID](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Process] [nvarchar](50) NOT NULL,
	[ProductIDList] [nvarchar](max) NULL,
 CONSTRAINT [PK_PRODUCT_ID] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PSA_PEEL_TEST_ON_PRODUCT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PSA_PEEL_TEST_ON_PRODUCT](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PSA_PEEL_TEST_ON_PRODUCT_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PSA_PEEL_TEST_ON_PRODUCT_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PSA_PEEL_TEST_ON_PRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PSA_PEEL_TEST_ON_PRODUCT_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_PSA_PEEL_TEST_ON_PRODUCT_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PT_ONPRODUCT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PT_ONPRODUCT](
	[Id] [int] NOT NULL,
	[ProductID] [nvarchar](50) NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Picture] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Peak(gf)] [nvarchar](20) NULL,
	[Average(gf)] [nvarchar](20) NULL,
	[Peak(N)] [nvarchar](20) NULL,
	[Average(N)] [nvarchar](20) NULL,
	[JudgementForce] [nvarchar](20) NULL,
	[JudgementMode] [nvarchar](50) NULL,
	[KeyDic] [nvarchar](50) NULL,
 CONSTRAINT [PK_PEEL_TEST_(ON_PRODUCT)] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PT_ONPRODUCT_BEFOREIMAGE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PT_ONPRODUCT_BEFOREIMAGE](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Picture] [varbinary](max) NULL,
 CONSTRAINT [PK_PEEL_TEST_(ON_PRODUCT)_BEFORE] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PT_ONPRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PT_ONPRODUCT_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[LocationImage] [nvarchar](200) NULL,
	[LocationSpec] [nvarchar](200) NULL,
	[LoactionBefore] [nvarchar](200) NULL,
	[Data] [nvarchar](max) NULL,
	[DataSpec] [nvarchar](max) NULL,
	[DataBefore] [nvarchar](max) NULL,
 CONSTRAINT [PK_PT_ONPRODUCT_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PT_ONPRODUCT_SEPC]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PT_ONPRODUCT_SEPC](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Type] [nvarchar](20) NULL,
	[Tape] [nvarchar](20) NULL,
	[Peak(N)] [nvarchar](20) NULL,
	[Average(N)] [nvarchar](20) NULL,
	[Peak(Gf)] [nvarchar](20) NULL,
	[Average(Gf)] [nvarchar](20) NULL,
 CONSTRAINT [PK_PEEL_TEST_(ON_PRODUCT)_SPEC] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Results]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Results](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](255) NULL,
	[LotNo] [nvarchar](255) NULL,
	[Items] [nvarchar](255) NULL,
	[Results] [nvarchar](255) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Roughness]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Roughness](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Machine] [nvarchar](max) NULL,
	[MDate] [nvarchar](max) NULL,
	[L1_Roughness_Sa] [nvarchar](max) NULL,
	[L1_Roughness_Sq] [nvarchar](max) NULL,
	[L1_Roughness_Sdr] [nvarchar](max) NULL,
	[L2_Roughness_Sa] [nvarchar](max) NULL,
	[L2_Roughness_Sq] [nvarchar](max) NULL,
	[L2_Roughness_Sdr] [nvarchar](max) NULL,
	[L3_Roughness_Sa] [nvarchar](max) NULL,
	[L3_Roughness_Sq] [nvarchar](max) NULL,
	[L3_Roughness_Sdr] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Roughness_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Roughness_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Data] [nvarchar](max) NULL,
 CONSTRAINT [PK_Roughness_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SEM_BSE_Binarization_Logfile]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SEM_BSE_Binarization_Logfile](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[SEM200250] [varbinary](max) NULL,
	[SEM500700] [varbinary](max) NULL,
	[SEM5K] [varbinary](max) NULL,
	[Judgement] [nvarchar](50) NULL,
	[Binarization200250] [varbinary](max) NULL,
	[Black200250] [float] NULL,
	[Binarization500700] [varbinary](max) NULL,
	[Black500700] [float] NULL,
	[CheckResults] [nchar](10) NULL,
 CONSTRAINT [PK_SEM_BSE_Binarization_Logfile] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SEM_BSE_Binarization_Logfile_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SEM_BSE_Binarization_Logfile_NAS](
	[Id] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Data] [nvarchar](max) NULL,
	[AddressImg] [nvarchar](max) NULL,
 CONSTRAINT [PK_SEM_BSE_Binarization_Logfile_NAS] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SHEAR_TEST]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SHEAR_TEST](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SHEAR_TEST_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SHEAR_TEST_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Sample] [nvarchar](max) NULL,
	[Image] [varbinary](max) NULL,
	[Graph] [varbinary](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Mode 1: Solder joint crack] [nvarchar](max) NULL,
	[Mode 2: Pad lift] [nvarchar](max) NULL,
	[Mode 3: Solder joint lift] [nvarchar](max) NULL,
	[Mode 4: Intermetallic break] [nvarchar](max) NULL,
	[Mode 5: Component damage] [nvarchar](max) NULL,
	[Mode 6: Component detached] [nvarchar](max) NULL,
	[Mode 7: Flex torn] [nvarchar](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SHEAR_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SHEAR_TEST_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Sheet] [nvarchar](50) NULL,
	[Data] [nvarchar](max) NULL,
	[LocationImg] [nvarchar](200) NULL,
 CONSTRAINT [PK_SHEAR_TEST_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SPEC_COMMENT_3]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SPEC_COMMENT_3](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[Sheet] [nvarchar](max) NULL,
	[Count_Sample] [nvarchar](max) NULL,
	[Location] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TABLE_OF_CONTENT]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TABLE_OF_CONTENT](
	[ID] [int] NOT NULL,
	[ItemCode] [char](20) NOT NULL,
	[LotNo] [char](20) NOT NULL,
	[BuildDate] [date] NULL,
	[SendDate] [date] NULL,
	[DeliveryQuatity] [char](20) NULL,
	[ShippingTo] [nvarchar](100) NULL,
	[ContentTable] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TABLE_OF_CONTENT_SETTING]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TABLE_OF_CONTENT_SETTING](
	[Id] [int] NOT NULL,
	[ItemName] [nchar](10) NULL,
	[ProgramName] [nvarchar](50) NULL,
	[ItemCode] [nchar](10) NULL,
	[MCORevision] [nvarchar](50) NULL,
	[ODBRevision] [nvarchar](50) NULL,
	[Build] [nvarchar](50) NULL,
	[XOUTRate] [nvarchar](50) NULL,
	[ShippingFrom] [nvarchar](50) NULL,
	[EEEECode] [nvarchar](50) NULL,
	[FactoryCode] [nvarchar](50) NULL,
 CONSTRAINT [PK_Table_1] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TARGET_OF_ASSY_YIELD]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TARGET_OF_ASSY_YIELD](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[Value] [nchar](10) NULL,
 CONSTRAINT [PK_TARGET_OF_ASSY_YIELD] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TC_HS_TS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TC_HS_TS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[DataLog] [nvarchar](max) NULL,
	[Type] [nchar](10) NULL,
 CONSTRAINT [PK_TC_HS_TS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[THERMAL_CYCLING_AND_BEND]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[THERMAL_CYCLING_AND_BEND](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Sel_report] [nvarchar](max) NULL,
	[Before] [nvarchar](max) NULL,
	[After_1] [nvarchar](max) NULL,
	[After_2] [nvarchar](max) NULL,
	[After_3] [nvarchar](max) NULL,
	[After_4] [nvarchar](max) NULL,
	[After_5] [nvarchar](max) NULL,
	[After_10] [nvarchar](max) NULL,
	[After_15] [nvarchar](max) NULL,
	[After_20] [nvarchar](max) NULL,
	[After_25] [nvarchar](max) NULL,
	[After_30] [nvarchar](max) NULL,
	[After_40] [nvarchar](max) NULL,
	[After_50] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[THERMAL_CYCLING_AND_BEND_LOGFILE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[THERMAL_CYCLING_AND_BEND_LOGFILE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Net_No] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Logfile] [nvarchar](max) NULL,
	[Cycles_name] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL,
	[Shift] [nvarchar](max) NULL,
	[UserID] [nvarchar](max) NULL,
	[ItemName] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TRACEWIDTH_IMAGE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TRACEWIDTH_IMAGE](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data_For] [nvarchar](max) NULL,
	[Image_Tracewidth] [varbinary](max) NULL,
	[Operator] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TRACEWIDTH_SPEC]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TRACEWIDTH_SPEC](
	[ID] [int] NOT NULL,
	[ItemCode] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Range] [nvarchar](max) NULL,
	[USL] [nvarchar](max) NULL,
	[LSL] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TRACEWIDTH_VAL]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TRACEWIDTH_VAL](
	[ID] [int] NULL,
	[ItemCode] [nvarchar](max) NULL,
	[LotNo] [nvarchar](max) NULL,
	[Region] [nvarchar](max) NULL,
	[Pcs_No] [nvarchar](max) NULL,
	[Data] [nvarchar](max) NULL,
	[Time_Update] [nvarchar](max) NULL,
	[Operaor] [nvarchar](max) NULL,
	[Data_For] [nvarchar](max) NULL,
	[Remark] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[XRAY]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[XRAY](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NOT NULL,
	[LotNo] [nchar](10) NOT NULL,
	[Area] [int] NULL,
	[Data] [nvarchar](max) NULL,
	[Type] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[XRAY_IMAGE]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[XRAY_IMAGE](
	[ID] [int] NOT NULL,
	[Area] [int] NOT NULL,
	[Image] [varbinary](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[XRAY_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[XRAY_NAS](
	[ID] [int] NOT NULL,
	[ItemCode] [nchar](10) NULL,
	[LotNo] [nchar](10) NULL,
	[Area] [nvarchar](200) NULL,
	[Data] [nvarchar](max) NULL,
	[Type] [nvarchar](50) NULL,
 CONSTRAINT [PK_XRAY_NAS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_ACCOUNT]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_ACCOUNT] ON [dbo].[ACCOUNT]
(
	[Username] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_ASSY_YIELD]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_ASSY_YIELD] ON [dbo].[ASSY_YIELD]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_CROSS_SECTION_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_CROSS_SECTION_NAS] ON [dbo].[CROSS_SECTION_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC,
	[Sheet] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_ENVIRONMENT_EN_DURANCE]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_ENVIRONMENT_EN_DURANCE] ON [dbo].[ENVIRONMENT_EN_DURANCE]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_FLEX_BENDING_S2]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE NONCLUSTERED INDEX [IX_FLEX_BENDING_S2] ON [dbo].[FLEX_BENDING_S2]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_GAP_CONNECTOR_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_GAP_CONNECTOR_NAS] ON [dbo].[GAP_CONNECTOR_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_LINER_PEEL_TEST_ON_PRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_LINER_PEEL_TEST_ON_PRODUCT_NAS] ON [dbo].[LINER_PEEL_TEST_ON_PRODUCT_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC,
	[Sheet] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_MATING_PULL_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_MATING_PULL_TEST_NAS] ON [dbo].[MATING_PULL_TEST_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_OQC_B2B_Mating_Unmating_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_OQC_B2B_Mating_Unmating_NAS] ON [dbo].[OQC_B2B_Mating_Unmating_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_PEEL_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_PEEL_TEST_NAS] ON [dbo].[PEEL_TEST_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_PEEL_TEST_WITHOUT_SUS_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_PEEL_TEST_WITHOUT_SUS_NAS] ON [dbo].[PEEL_TEST_WITHOUT_SUS_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC,
	[Sheet] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_PSA_PEEL_TEST_ON_PRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_PSA_PEEL_TEST_ON_PRODUCT_NAS] ON [dbo].[PSA_PEEL_TEST_ON_PRODUCT_NAS]
(
	[ItemCode] DESC,
	[LotNo] DESC,
	[Sheet] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_PT_ONPRODUCT_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_PT_ONPRODUCT_NAS] ON [dbo].[PT_ONPRODUCT_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_SEM_BSE_Binarization_Logfile]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE NONCLUSTERED INDEX [IX_SEM_BSE_Binarization_Logfile] ON [dbo].[SEM_BSE_Binarization_Logfile]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_SEM_BSE_Binarization_Logfile_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_SEM_BSE_Binarization_Logfile_NAS] ON [dbo].[SEM_BSE_Binarization_Logfile_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_SHEAR_TEST_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_SHEAR_TEST_NAS] ON [dbo].[SHEAR_TEST_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC,
	[Sheet] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_TC_HS_TS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE UNIQUE NONCLUSTERED INDEX [IX_TC_HS_TS] ON [dbo].[TC_HS_TS]
(
	[ItemCode] ASC,
	[LotNo] ASC,
	[Type] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_XRAY_NAS]    Script Date: 3/20/2026 5:31:45 PM ******/
CREATE NONCLUSTERED INDEX [IX_XRAY_NAS] ON [dbo].[XRAY_NAS]
(
	[ItemCode] ASC,
	[LotNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
USE [master]
GO
ALTER DATABASE [OK2SHIP_SMT] SET  READ_WRITE 
GO
