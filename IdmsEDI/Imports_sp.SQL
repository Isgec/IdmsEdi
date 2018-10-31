USE [inforerpdb]
GO

/****** Object:  StoredProcedure [dbo].[spediAFileDelete]    Script Date: 15/06/2018 18:26:14 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediAFileDelete]
  @Original_t_drid VarChar(50),
  @RowCount int = null OUTPUT
  AS
  DELETE [ttcisg132200]
  WHERE
  [ttcisg132200].[t_drid] = @Original_t_drid
  SET @RowCount = @@RowCount
GO

/****** Object:  StoredProcedure [dbo].[spediAFileUpdate]    Script Date: 15/06/2018 18:26:14 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediAFileUpdate]
  @Original_t_drid VarChar(50), 
  @t_dcid VarChar(200),
  @t_hndl VarChar(200),
  @t_indx VarChar(50),
  @t_prcd VarChar(50),
  @t_fnam VarChar(250),
  @t_lbcd VarChar(50),
  @t_atby VarChar(50),
  @t_aton DateTime,
  @t_Refcntd Int,
  @t_Refcntu Int,
  @RowCount int = null OUTPUT
  AS
  UPDATE [ttcisg132200] SET 
   [t_dcid] = @t_dcid
  ,[t_hndl] = @t_hndl
  ,[t_indx] = @t_indx
  ,[t_prcd] = @t_prcd
  ,[t_fnam] = @t_fnam
  ,[t_lbcd] = @t_lbcd
  ,[t_atby] = @t_atby
  ,[t_aton] = @t_aton
  ,[t_Refcntd] = @t_Refcntd
  ,[t_Refcntu] = @t_Refcntu
  WHERE
  [t_drid] = @Original_t_drid
  SET @RowCount = @@RowCount
GO

/****** Object:  StoredProcedure [dbo].[spediAFileInsert]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediAFileInsert]
  @t_dcid VarChar(200),
  @t_hndl VarChar(200),
  @t_indx VarChar(50),
  @t_prcd VarChar(50),
  @t_fnam VarChar(250),
  @t_lbcd VarChar(50),
  @t_atby VarChar(50),
  @t_aton DateTime,
  @t_Refcntd Int,
  @t_Refcntu Int,
  @Return_t_drid VarChar(50) = null OUTPUT 
  AS
  BEGIN 
  Declare @drid bigint
  select @drid = max(convert(bigint,t_drid)) from ttcisg132200 
  
  INSERT [ttcisg132200]
  (
   [t_drid]
  ,[t_dcid]
  ,[t_hndl]
  ,[t_indx]
  ,[t_prcd]
  ,[t_fnam]
  ,[t_lbcd]
  ,[t_atby]
  ,[t_aton]
  ,[t_Refcntd]
  ,[t_Refcntu]
  )
  VALUES
  (
   @drid + 1
  ,@t_dcid
  ,@t_hndl
  ,@t_indx
  ,@t_prcd
  ,@t_fnam
  ,@t_lbcd
  ,@t_atby
  ,@t_aton
  ,@t_Refcntd
  ,@t_Refcntu
  )
  SET @Return_t_drid = @drid
   
 END
GO

/****** Object:  StoredProcedure [dbo].[spediAFileSelectListFilteres]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediAFileSelectListFilteres]
  @Filter_t_drid VarChar(50),
  @Filter_t_hndl VarChar(200),
  @Filter_t_indx VarChar(50),
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
  DECLARE @LGSQL VarChar(8000)
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_drid VarChar(50) NOT NULL
  )
  SET @LGSQL = 'INSERT INTO #PageIndex (' 
  SET @LGSQL = @LGSQL + 't_drid'
  SET @LGSQL = @LGSQL + ')'
  SET @LGSQL = @LGSQL + ' SELECT '
  SET @LGSQL = @LGSQL + '[ttcisg132200].[t_drid]'
  SET @LGSQL = @LGSQL + ' FROM [ttcisg132200] '
  SET @LGSQL = @LGSQL + '  WHERE 1 = 1 '
  IF (@Filter_t_drid > '') 
    SET @LGSQL = @LGSQL + ' AND [ttcisg132200].[t_drid] = ''' + @Filter_t_drid + ''''
  IF (@Filter_t_hndl > '') 
    SET @LGSQL = @LGSQL + ' AND [ttcisg132200].[t_hndl] = ''' + @Filter_t_hndl + ''''
  IF (@Filter_t_indx > '') 
    SET @LGSQL = @LGSQL + ' AND [ttcisg132200].[t_indx] = ''' + @Filter_t_indx + ''''
  SET @LGSQL = @LGSQL + '  ORDER BY '
  SET @LGSQL = @LGSQL + CASE @OrderBy
                        WHEN 't_drid' THEN '[ttcisg132200].[t_drid]'
                        WHEN 't_drid DESC' THEN '[ttcisg132200].[t_drid] DESC'
                        WHEN 't_dcid' THEN '[ttcisg132200].[t_dcid]'
                        WHEN 't_dcid DESC' THEN '[ttcisg132200].[t_dcid] DESC'
                        WHEN 't_hndl' THEN '[ttcisg132200].[t_hndl]'
                        WHEN 't_hndl DESC' THEN '[ttcisg132200].[t_hndl] DESC'
                        WHEN 't_indx' THEN '[ttcisg132200].[t_indx]'
                        WHEN 't_indx DESC' THEN '[ttcisg132200].[t_indx] DESC'
                        WHEN 't_prcd' THEN '[ttcisg132200].[t_prcd]'
                        WHEN 't_prcd DESC' THEN '[ttcisg132200].[t_prcd] DESC'
                        WHEN 't_fnam' THEN '[ttcisg132200].[t_fnam]'
                        WHEN 't_fnam DESC' THEN '[ttcisg132200].[t_fnam] DESC'
                        WHEN 't_lbcd' THEN '[ttcisg132200].[t_lbcd]'
                        WHEN 't_lbcd DESC' THEN '[ttcisg132200].[t_lbcd] DESC'
                        WHEN 't_atby' THEN '[ttcisg132200].[t_atby]'
                        WHEN 't_atby DESC' THEN '[ttcisg132200].[t_atby] DESC'
                        WHEN 't_aton' THEN '[ttcisg132200].[t_aton]'
                        WHEN 't_aton DESC' THEN '[ttcisg132200].[t_aton] DESC'
                        WHEN 't_Refcntd' THEN '[ttcisg132200].[t_Refcntd]'
                        WHEN 't_Refcntd DESC' THEN '[ttcisg132200].[t_Refcntd] DESC'
                        WHEN 't_Refcntu' THEN '[ttcisg132200].[t_Refcntu]'
                        WHEN 't_Refcntu DESC' THEN '[ttcisg132200].[t_Refcntu] DESC'
                        ELSE '[ttcisg132200].[t_drid]'
                    END
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [ttcisg132200].*  
  FROM [ttcisg132200] 
      INNER JOIN #PageIndex
          ON [ttcisg132200].[t_drid] = #PageIndex.t_drid
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
GO

/****** Object:  StoredProcedure [dbo].[spediAFileSelectByID]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediAFileSelectByID]
  @LoginID NVarChar(8),
  @t_drid VarChar(50) 
  AS
  SELECT
    [ttcisg132200].*  
  FROM [ttcisg132200] 
  WHERE
  [ttcisg132200].[t_drid] = @t_drid
GO

/****** Object:  StoredProcedure [dbo].[spediASeriesUpdate]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediASeriesUpdate]
  @Original_t_rnum Int, 
  @t_seri VarChar(50),
  @t_rnum Int,
  @t_acti VarChar(1),
  @t_Refcntd Int,
  @t_Refcntu Int,
  @RowCount int = null OUTPUT
  AS
  UPDATE [ttcisg131200] SET 
   [t_seri] = @t_seri
  ,[t_rnum] = @t_rnum
  ,[t_acti] = @t_acti
  ,[t_Refcntd] = @t_Refcntd
  ,[t_Refcntu] = @t_Refcntu
  WHERE
  [t_rnum] = @Original_t_rnum
  SET @RowCount = @@RowCount
GO

/****** Object:  StoredProcedure [dbo].[spediASeriesSelectListFilteres]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediASeriesSelectListFilteres]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @t_acti VarChar(1),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
  DECLARE @LGSQL VarChar(8000)
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_rnum Int NOT NULL
  )
  SET @LGSQL = 'INSERT INTO #PageIndex (' 
  SET @LGSQL = @LGSQL + 't_rnum'
  SET @LGSQL = @LGSQL + ')'
  SET @LGSQL = @LGSQL + ' SELECT '
  SET @LGSQL = @LGSQL + '[ttcisg131200].[t_rnum]'
  SET @LGSQL = @LGSQL + ' FROM [ttcisg131200] '
  SET @LGSQL = @LGSQL + '  WHERE 1 = 1 '
  SET @LGSQL = @LGSQL + ' AND [ttcisg131200].[t_acti] = (''' + @t_acti + ''')'
  SET @LGSQL = @LGSQL + '  ORDER BY '
  SET @LGSQL = @LGSQL + CASE @OrderBy
                        WHEN 't_seri' THEN '[ttcisg131200].[t_seri]'
                        WHEN 't_seri DESC' THEN '[ttcisg131200].[t_seri] DESC'
                        WHEN 't_rnum' THEN '[ttcisg131200].[t_rnum]'
                        WHEN 't_rnum DESC' THEN '[ttcisg131200].[t_rnum] DESC'
                        WHEN 't_acti' THEN '[ttcisg131200].[t_acti]'
                        WHEN 't_acti DESC' THEN '[ttcisg131200].[t_acti] DESC'
                        WHEN 't_Refcntd' THEN '[ttcisg131200].[t_Refcntd]'
                        WHEN 't_Refcntd DESC' THEN '[ttcisg131200].[t_Refcntd] DESC'
                        WHEN 't_Refcntu' THEN '[ttcisg131200].[t_Refcntu]'
                        WHEN 't_Refcntu DESC' THEN '[ttcisg131200].[t_Refcntu] DESC'
                        ELSE '[ttcisg131200].[t_rnum]'
                    END
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [ttcisg131200].*  
  FROM [ttcisg131200] 
      INNER JOIN #PageIndex
          ON [ttcisg131200].[t_rnum] = #PageIndex.t_rnum
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
GO

/****** Object:  StoredProcedure [dbo].[spediASeriesSelectByID]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediASeriesSelectByID]
  @LoginID NVarChar(8),
  @t_rnum Int 
  AS
  SELECT
    [ttcisg131200].*  
  FROM [ttcisg131200] 
  WHERE
  [ttcisg131200].[t_rnum] = @t_rnum
GO

/****** Object:  StoredProcedure [dbo].[spediALibSelectListFilteres]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediALibSelectListFilteres]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @t_acti VarChar(1),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
  DECLARE @LGSQL VarChar(8000)
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_lbcd VarChar(50) NOT NULL
  )
  SET @LGSQL = 'INSERT INTO #PageIndex (' 
  SET @LGSQL = @LGSQL + 't_lbcd'
  SET @LGSQL = @LGSQL + ')'
  SET @LGSQL = @LGSQL + ' SELECT '
  SET @LGSQL = @LGSQL + '[ttcisg127200].[t_lbcd]'
  SET @LGSQL = @LGSQL + ' FROM [ttcisg127200] '
  SET @LGSQL = @LGSQL + '  WHERE 1 = 1 '
  SET @LGSQL = @LGSQL + ' AND [ttcisg127200].[t_acti] = (''' + @t_acti + ''')'
  SET @LGSQL = @LGSQL + '  ORDER BY '
  SET @LGSQL = @LGSQL + CASE @OrderBy
                        WHEN 't_lbcd' THEN '[ttcisg127200].[t_lbcd]'
                        WHEN 't_lbcd DESC' THEN '[ttcisg127200].[t_lbcd] DESC'
                        WHEN 't_desc' THEN '[ttcisg127200].[t_desc]'
                        WHEN 't_desc DESC' THEN '[ttcisg127200].[t_desc] DESC'
                        WHEN 't_serv' THEN '[ttcisg127200].[t_serv]'
                        WHEN 't_serv DESC' THEN '[ttcisg127200].[t_serv] DESC'
                        WHEN 't_path' THEN '[ttcisg127200].[t_path]'
                        WHEN 't_path DESC' THEN '[ttcisg127200].[t_path] DESC'
                        WHEN 't_acti' THEN '[ttcisg127200].[t_acti]'
                        WHEN 't_acti DESC' THEN '[ttcisg127200].[t_acti] DESC'
                        WHEN 't_Refcntd' THEN '[ttcisg127200].[t_Refcntd]'
                        WHEN 't_Refcntd DESC' THEN '[ttcisg127200].[t_Refcntd] DESC'
                        WHEN 't_Refcntu' THEN '[ttcisg127200].[t_Refcntu]'
                        WHEN 't_Refcntu DESC' THEN '[ttcisg127200].[t_Refcntu] DESC'
                        ELSE '[ttcisg127200].[t_lbcd]'
                    END
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [ttcisg127200].*  
  FROM [ttcisg127200] 
      INNER JOIN #PageIndex
          ON [ttcisg127200].[t_lbcd] = #PageIndex.t_lbcd
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
GO

/****** Object:  StoredProcedure [dbo].[spediALibSelectByID]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediALibSelectByID]
  @LoginID NVarChar(8),
  @t_lbcd VarChar(50) 
  AS
  SELECT
    [ttcisg127200].*  
  FROM [ttcisg127200] 
  WHERE
  [ttcisg127200].[t_lbcd] = @t_lbcd
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlDUpdate]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlDUpdate]
  @Original_t_tran VarChar(9), 
  @Original_t_docn VarChar(32), 
  @Original_t_revn VarChar(20), 
  @t_tran VarChar(9),
  @t_docn VarChar(32),
  @t_revn VarChar(20),
  @t_stid VarChar(3),
  @t_pono Int,
  @t_remk VarChar(100),
  @t_recv VarChar(3),
  @t_refr VarChar(30),
  @t_redt DateTime,
  @t_rekm VarChar(100),
  @t_lock Int,
  @t_recc Int,
  @t_revd Int,
  @t_issu Int,
  @t_Refcntd Int,
  @t_Refcntu Int,
  @RowCount int = null OUTPUT
  AS
  UPDATE [tdmisg132200] SET 
   [t_tran] = @t_tran
  ,[t_docn] = @t_docn
  ,[t_revn] = @t_revn
  ,[t_stid] = @t_stid
  ,[t_pono] = @t_pono
  ,[t_remk] = @t_remk
  ,[t_recv] = @t_recv
  ,[t_refr] = @t_refr
  ,[t_redt] = @t_redt
  ,[t_rekm] = @t_rekm
  ,[t_lock] = @t_lock
  ,[t_recc] = @t_recc
  ,[t_revd] = @t_revd
  ,[t_issu] = @t_issu
  ,[t_Refcntd] = @t_Refcntd
  ,[t_Refcntu] = @t_Refcntu
  WHERE
  [t_tran] = @Original_t_tran
  AND [t_docn] = @Original_t_docn
  AND [t_revn] = @Original_t_revn
  SET @RowCount = @@RowCount
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlDSelectListFilteres]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlDSelectListFilteres]
  @Filter_t_tran VarChar(9),
  @Filter_t_docn VarChar(32),
  @Filter_t_revn VarChar(20),
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
  DECLARE @LGSQL VarChar(8000)
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_tran VarChar(9) NOT NULL
 ,t_docn VarChar(32) NOT NULL
 ,t_revn VarChar(20) NOT NULL
  )
  SET @LGSQL = 'INSERT INTO #PageIndex (' 
  SET @LGSQL = @LGSQL + 't_tran'
  SET @LGSQL = @LGSQL + ', t_docn'
  SET @LGSQL = @LGSQL + ', t_revn'
  SET @LGSQL = @LGSQL + ')'
  SET @LGSQL = @LGSQL + ' SELECT '
  SET @LGSQL = @LGSQL + '[tdmisg132200].[t_tran]'
  SET @LGSQL = @LGSQL + ', [tdmisg132200].[t_docn]'
  SET @LGSQL = @LGSQL + ', [tdmisg132200].[t_revn]'
  SET @LGSQL = @LGSQL + ' FROM [tdmisg132200] '
  SET @LGSQL = @LGSQL + '  WHERE 1 = 1 '
  IF (@Filter_t_tran > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg132200].[t_tran] = ''' + @Filter_t_tran + ''''
  IF (@Filter_t_docn > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg132200].[t_docn] = ''' + @Filter_t_docn + ' COLLATE LATIN1_GENERAL_BIN2'''
  IF (@Filter_t_revn > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg132200].[t_revn] = ''' + @Filter_t_revn + ' COLLATE LATIN1_GENERAL_BIN2'''
  SET @LGSQL = @LGSQL + '  ORDER BY '
  SET @LGSQL = @LGSQL + CASE @OrderBy
                        WHEN 't_tran' THEN '[tdmisg132200].[t_tran]'
                        WHEN 't_tran DESC' THEN '[tdmisg132200].[t_tran] DESC'
                        WHEN 't_docn' THEN '[tdmisg132200].[t_docn]'
                        WHEN 't_docn DESC' THEN '[tdmisg132200].[t_docn] DESC'
                        WHEN 't_revn' THEN '[tdmisg132200].[t_revn]'
                        WHEN 't_revn DESC' THEN '[tdmisg132200].[t_revn] DESC'
                        WHEN 't_stid' THEN '[tdmisg132200].[t_stid]'
                        WHEN 't_stid DESC' THEN '[tdmisg132200].[t_stid] DESC'
                        WHEN 't_pono' THEN '[tdmisg132200].[t_pono]'
                        WHEN 't_pono DESC' THEN '[tdmisg132200].[t_pono] DESC'
                        WHEN 't_remk' THEN '[tdmisg132200].[t_remk]'
                        WHEN 't_remk DESC' THEN '[tdmisg132200].[t_remk] DESC'
                        WHEN 't_recv' THEN '[tdmisg132200].[t_recv]'
                        WHEN 't_recv DESC' THEN '[tdmisg132200].[t_recv] DESC'
                        WHEN 't_refr' THEN '[tdmisg132200].[t_refr]'
                        WHEN 't_refr DESC' THEN '[tdmisg132200].[t_refr] DESC'
                        WHEN 't_redt' THEN '[tdmisg132200].[t_redt]'
                        WHEN 't_redt DESC' THEN '[tdmisg132200].[t_redt] DESC'
                        WHEN 't_rekm' THEN '[tdmisg132200].[t_rekm]'
                        WHEN 't_rekm DESC' THEN '[tdmisg132200].[t_rekm] DESC'
                        WHEN 't_lock' THEN '[tdmisg132200].[t_lock]'
                        WHEN 't_lock DESC' THEN '[tdmisg132200].[t_lock] DESC'
                        WHEN 't_recc' THEN '[tdmisg132200].[t_recc]'
                        WHEN 't_recc DESC' THEN '[tdmisg132200].[t_recc] DESC'
                        WHEN 't_revd' THEN '[tdmisg132200].[t_revd]'
                        WHEN 't_revd DESC' THEN '[tdmisg132200].[t_revd] DESC'
                        WHEN 't_issu' THEN '[tdmisg132200].[t_issu]'
                        WHEN 't_issu DESC' THEN '[tdmisg132200].[t_issu] DESC'
                        WHEN 't_Refcntd' THEN '[tdmisg132200].[t_Refcntd]'
                        WHEN 't_Refcntd DESC' THEN '[tdmisg132200].[t_Refcntd] DESC'
                        WHEN 't_Refcntu' THEN '[tdmisg132200].[t_Refcntu]'
                        WHEN 't_Refcntu DESC' THEN '[tdmisg132200].[t_Refcntu] DESC'
                        ELSE '[tdmisg132200].[t_tran],[tdmisg132200].[t_docn],[tdmisg132200].[t_revn]'
                    END
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [tdmisg132200].*  
  FROM [tdmisg132200] 
      INNER JOIN #PageIndex
          ON [tdmisg132200].[t_tran] = #PageIndex.t_tran COLLATE LATIN1_GENERAL_BIN2
          AND [tdmisg132200].[t_docn] = #PageIndex.t_docn COLLATE LATIN1_GENERAL_BIN2
          AND [tdmisg132200].[t_revn] = #PageIndex.t_revn COLLATE LATIN1_GENERAL_BIN2
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlDSelectByID]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlDSelectByID]
  @LoginID NVarChar(8),
  @t_tran VarChar(9),
  @t_docn VarChar(32),
  @t_revn VarChar(20) 
  AS
  SELECT
    [tdmisg132200].*  
  FROM [tdmisg132200] 
  WHERE
  [tdmisg132200].[t_tran] = @t_tran
  AND [tdmisg132200].[t_docn] = @t_docn
  AND [tdmisg132200].[t_revn] = @t_revn
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlHUpdate]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlHUpdate]
  @Original_t_tran VarChar(9), 
  @t_tran VarChar(9),
  @t_type Int,
  @t_bpid VarChar(9),
  @t_cadr VarChar(9),
  @t_cprj VarChar(9),
  @t_logn VarChar(16),
  @t_subj VarChar(100),
  @t_remk VarChar(100),
  @t_issu VarChar(3),
  @t_stat Int,
  @t_ofbp VarChar(9),
  @t_vadr VarChar(9),
  @t_padr VarChar(9),
  @t_dprj VarChar(9),
  @t_user VarChar(16),
  @t_date DateTime,
  @t_subt Int,
  @t_refr VarChar(32),
  @t_appr Int,
  @t_rejc Int,
  @t_rekm VarChar(100),
  @t_apdt DateTime,
  @t_apsu VarChar(16),
  @t_irmk VarChar(100),
  @t_iisu Int,
  @t_retn Int,
  @t_smdt DateTime,
  @t_isby VarChar(16),
  @t_isdt DateTime,
  @t_Refcntd Int,
  @t_Refcntu Int,
  @t_edif Int,
  @RowCount int = null OUTPUT
  AS
  UPDATE [tdmisg131200] SET 
   [t_tran] = @t_tran
  ,[t_type] = @t_type
  ,[t_bpid] = @t_bpid
  ,[t_cadr] = @t_cadr
  ,[t_cprj] = @t_cprj
  ,[t_logn] = @t_logn
  ,[t_subj] = @t_subj
  ,[t_remk] = @t_remk
  ,[t_issu] = @t_issu
  ,[t_stat] = @t_stat
  ,[t_ofbp] = @t_ofbp
  ,[t_vadr] = @t_vadr
  ,[t_padr] = @t_padr
  ,[t_dprj] = @t_dprj
  ,[t_user] = @t_user
  ,[t_date] = @t_date
  ,[t_subt] = @t_subt
  ,[t_refr] = @t_refr
  ,[t_appr] = @t_appr
  ,[t_rejc] = @t_rejc
  ,[t_rekm] = @t_rekm
  ,[t_apdt] = @t_apdt
  ,[t_apsu] = @t_apsu
  ,[t_irmk] = @t_irmk
  ,[t_iisu] = @t_iisu
  ,[t_retn] = @t_retn
  ,[t_smdt] = @t_smdt
  ,[t_isby] = @t_isby
  ,[t_isdt] = @t_isdt
  ,[t_Refcntd] = @t_Refcntd
  ,[t_Refcntu] = @t_Refcntu
  ,[t_edif] = @t_edif
  WHERE
  [t_tran] = @Original_t_tran
  SET @RowCount = @@RowCount
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlHSelectListFilteres]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlHSelectListFilteres]
  @Filter_t_tran VarChar(9),
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
  DECLARE @LGSQL VarChar(8000)
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_tran VarChar(9) NOT NULL
  )
  SET @LGSQL = 'INSERT INTO #PageIndex (' 
  SET @LGSQL = @LGSQL + 't_tran'
  SET @LGSQL = @LGSQL + ')'
  SET @LGSQL = @LGSQL + ' SELECT '
  SET @LGSQL = @LGSQL + '[tdmisg131200].[t_tran]'
  SET @LGSQL = @LGSQL + ' FROM [tdmisg131200] '
  SET @LGSQL = @LGSQL + '  WHERE [tdmisg131200].[t_edif] = 2 AND [tdmisg131200].[t_stat] = 5  and [tdmisg131200].[t_isdt] >= convert(datetime,''15/06/2018'',103)'
  IF (@Filter_t_tran > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg131200].[t_tran] = ''' + @Filter_t_tran + ''''
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [tdmisg131200].*  
  FROM [tdmisg131200] 
      INNER JOIN #PageIndex
          ON [tdmisg131200].[t_tran] = #PageIndex.t_tran COLLATE LATIN1_GENERAL_BIN2
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
GO

/****** Object:  StoredProcedure [dbo].[spediTmtlHSelectByID]    Script Date: 15/06/2018 18:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spediTmtlHSelectByID]
  @LoginID NVarChar(8),
  @t_tran VarChar(9) 
  AS
  SELECT
    [tdmisg131200].*  
  FROM [tdmisg131200] 
  WHERE
  [tdmisg131200].[t_tran] = @t_tran
GO


