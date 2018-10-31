SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHSelectByID]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHSelectByID]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHSelectListSearch]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHSelectListSearch]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlHSelectListSearch]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @KeyWord VarChar(250),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
    DECLARE @KeyWord1 VarChar(260)
    SET @KeyWord1 = '%' + LOWER(@KeyWord) + '%'
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_tran VarChar(9) NOT NULL
  )
  INSERT INTO #PageIndex (t_tran)
  SELECT [tdmisg131200].[t_tran] FROM [tdmisg131200]
 WHERE  
   ( 
         LOWER(ISNULL([tdmisg131200].[t_tran],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_type], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_bpid],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_cadr],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_cprj],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_logn],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_subj],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_remk],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_issu],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_stat], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_ofbp],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_vadr],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_padr],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_dprj],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_user],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_subt], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_refr],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_appr], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_rejc], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_rekm],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_apsu],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_irmk],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_iisu], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_retn], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg131200].[t_isby],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_Refcntd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_Refcntu], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg131200].[t_edif], 0)) LIKE @KeyWord1
   ) 
  ORDER BY
     CASE @OrderBy WHEN 't_tran' THEN [tdmisg131200].[t_tran] END,
     CASE @OrderBy WHEN 't_tran DESC' THEN [tdmisg131200].[t_tran] END DESC,
     CASE @OrderBy WHEN 't_type' THEN [tdmisg131200].[t_type] END,
     CASE @OrderBy WHEN 't_type DESC' THEN [tdmisg131200].[t_type] END DESC,
     CASE @OrderBy WHEN 't_bpid' THEN [tdmisg131200].[t_bpid] END,
     CASE @OrderBy WHEN 't_bpid DESC' THEN [tdmisg131200].[t_bpid] END DESC,
     CASE @OrderBy WHEN 't_cadr' THEN [tdmisg131200].[t_cadr] END,
     CASE @OrderBy WHEN 't_cadr DESC' THEN [tdmisg131200].[t_cadr] END DESC,
     CASE @OrderBy WHEN 't_cprj' THEN [tdmisg131200].[t_cprj] END,
     CASE @OrderBy WHEN 't_cprj DESC' THEN [tdmisg131200].[t_cprj] END DESC,
     CASE @OrderBy WHEN 't_logn' THEN [tdmisg131200].[t_logn] END,
     CASE @OrderBy WHEN 't_logn DESC' THEN [tdmisg131200].[t_logn] END DESC,
     CASE @OrderBy WHEN 't_subj' THEN [tdmisg131200].[t_subj] END,
     CASE @OrderBy WHEN 't_subj DESC' THEN [tdmisg131200].[t_subj] END DESC,
     CASE @OrderBy WHEN 't_remk' THEN [tdmisg131200].[t_remk] END,
     CASE @OrderBy WHEN 't_remk DESC' THEN [tdmisg131200].[t_remk] END DESC,
     CASE @OrderBy WHEN 't_issu' THEN [tdmisg131200].[t_issu] END,
     CASE @OrderBy WHEN 't_issu DESC' THEN [tdmisg131200].[t_issu] END DESC,
     CASE @OrderBy WHEN 't_stat' THEN [tdmisg131200].[t_stat] END,
     CASE @OrderBy WHEN 't_stat DESC' THEN [tdmisg131200].[t_stat] END DESC,
     CASE @OrderBy WHEN 't_ofbp' THEN [tdmisg131200].[t_ofbp] END,
     CASE @OrderBy WHEN 't_ofbp DESC' THEN [tdmisg131200].[t_ofbp] END DESC,
     CASE @OrderBy WHEN 't_vadr' THEN [tdmisg131200].[t_vadr] END,
     CASE @OrderBy WHEN 't_vadr DESC' THEN [tdmisg131200].[t_vadr] END DESC,
     CASE @OrderBy WHEN 't_padr' THEN [tdmisg131200].[t_padr] END,
     CASE @OrderBy WHEN 't_padr DESC' THEN [tdmisg131200].[t_padr] END DESC,
     CASE @OrderBy WHEN 't_dprj' THEN [tdmisg131200].[t_dprj] END,
     CASE @OrderBy WHEN 't_dprj DESC' THEN [tdmisg131200].[t_dprj] END DESC,
     CASE @OrderBy WHEN 't_user' THEN [tdmisg131200].[t_user] END,
     CASE @OrderBy WHEN 't_user DESC' THEN [tdmisg131200].[t_user] END DESC,
     CASE @OrderBy WHEN 't_date' THEN [tdmisg131200].[t_date] END,
     CASE @OrderBy WHEN 't_date DESC' THEN [tdmisg131200].[t_date] END DESC,
     CASE @OrderBy WHEN 't_subt' THEN [tdmisg131200].[t_subt] END,
     CASE @OrderBy WHEN 't_subt DESC' THEN [tdmisg131200].[t_subt] END DESC,
     CASE @OrderBy WHEN 't_refr' THEN [tdmisg131200].[t_refr] END,
     CASE @OrderBy WHEN 't_refr DESC' THEN [tdmisg131200].[t_refr] END DESC,
     CASE @OrderBy WHEN 't_appr' THEN [tdmisg131200].[t_appr] END,
     CASE @OrderBy WHEN 't_appr DESC' THEN [tdmisg131200].[t_appr] END DESC,
     CASE @OrderBy WHEN 't_rejc' THEN [tdmisg131200].[t_rejc] END,
     CASE @OrderBy WHEN 't_rejc DESC' THEN [tdmisg131200].[t_rejc] END DESC,
     CASE @OrderBy WHEN 't_rekm' THEN [tdmisg131200].[t_rekm] END,
     CASE @OrderBy WHEN 't_rekm DESC' THEN [tdmisg131200].[t_rekm] END DESC,
     CASE @OrderBy WHEN 't_apdt' THEN [tdmisg131200].[t_apdt] END,
     CASE @OrderBy WHEN 't_apdt DESC' THEN [tdmisg131200].[t_apdt] END DESC,
     CASE @OrderBy WHEN 't_apsu' THEN [tdmisg131200].[t_apsu] END,
     CASE @OrderBy WHEN 't_apsu DESC' THEN [tdmisg131200].[t_apsu] END DESC,
     CASE @OrderBy WHEN 't_irmk' THEN [tdmisg131200].[t_irmk] END,
     CASE @OrderBy WHEN 't_irmk DESC' THEN [tdmisg131200].[t_irmk] END DESC,
     CASE @OrderBy WHEN 't_iisu' THEN [tdmisg131200].[t_iisu] END,
     CASE @OrderBy WHEN 't_iisu DESC' THEN [tdmisg131200].[t_iisu] END DESC,
     CASE @OrderBy WHEN 't_retn' THEN [tdmisg131200].[t_retn] END,
     CASE @OrderBy WHEN 't_retn DESC' THEN [tdmisg131200].[t_retn] END DESC,
     CASE @OrderBy WHEN 't_smdt' THEN [tdmisg131200].[t_smdt] END,
     CASE @OrderBy WHEN 't_smdt DESC' THEN [tdmisg131200].[t_smdt] END DESC,
     CASE @OrderBy WHEN 't_isby' THEN [tdmisg131200].[t_isby] END,
     CASE @OrderBy WHEN 't_isby DESC' THEN [tdmisg131200].[t_isby] END DESC,
     CASE @OrderBy WHEN 't_isdt' THEN [tdmisg131200].[t_isdt] END,
     CASE @OrderBy WHEN 't_isdt DESC' THEN [tdmisg131200].[t_isdt] END DESC,
     CASE @OrderBy WHEN 't_Refcntd' THEN [tdmisg131200].[t_Refcntd] END,
     CASE @OrderBy WHEN 't_Refcntd DESC' THEN [tdmisg131200].[t_Refcntd] END DESC,
     CASE @OrderBy WHEN 't_Refcntu' THEN [tdmisg131200].[t_Refcntu] END,
     CASE @OrderBy WHEN 't_Refcntu DESC' THEN [tdmisg131200].[t_Refcntu] END DESC,
     CASE @OrderBy WHEN 't_edif' THEN [tdmisg131200].[t_edif] END,
     CASE @OrderBy WHEN 't_edif DESC' THEN [tdmisg131200].[t_edif] END DESC 

    SET @RecordCount = @@RowCount

  SELECT
    [tdmisg131200].*  
  FROM [tdmisg131200] 
      INNER JOIN #PageIndex
          ON [tdmisg131200].[t_tran] = #PageIndex.t_tran
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHSelectListFilteres]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHSelectListFilteres]
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
  SET @LGSQL = @LGSQL + '  WHERE 1 = 1 '
  IF (@Filter_t_tran > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg131200].[t_tran] = ''' + @Filter_t_tran + ''''
  SET @LGSQL = @LGSQL + '  ORDER BY '
  SET @LGSQL = @LGSQL + CASE @OrderBy
                        WHEN 't_tran' THEN '[tdmisg131200].[t_tran]'
                        WHEN 't_tran DESC' THEN '[tdmisg131200].[t_tran] DESC'
                        WHEN 't_type' THEN '[tdmisg131200].[t_type]'
                        WHEN 't_type DESC' THEN '[tdmisg131200].[t_type] DESC'
                        WHEN 't_bpid' THEN '[tdmisg131200].[t_bpid]'
                        WHEN 't_bpid DESC' THEN '[tdmisg131200].[t_bpid] DESC'
                        WHEN 't_cadr' THEN '[tdmisg131200].[t_cadr]'
                        WHEN 't_cadr DESC' THEN '[tdmisg131200].[t_cadr] DESC'
                        WHEN 't_cprj' THEN '[tdmisg131200].[t_cprj]'
                        WHEN 't_cprj DESC' THEN '[tdmisg131200].[t_cprj] DESC'
                        WHEN 't_logn' THEN '[tdmisg131200].[t_logn]'
                        WHEN 't_logn DESC' THEN '[tdmisg131200].[t_logn] DESC'
                        WHEN 't_subj' THEN '[tdmisg131200].[t_subj]'
                        WHEN 't_subj DESC' THEN '[tdmisg131200].[t_subj] DESC'
                        WHEN 't_remk' THEN '[tdmisg131200].[t_remk]'
                        WHEN 't_remk DESC' THEN '[tdmisg131200].[t_remk] DESC'
                        WHEN 't_issu' THEN '[tdmisg131200].[t_issu]'
                        WHEN 't_issu DESC' THEN '[tdmisg131200].[t_issu] DESC'
                        WHEN 't_stat' THEN '[tdmisg131200].[t_stat]'
                        WHEN 't_stat DESC' THEN '[tdmisg131200].[t_stat] DESC'
                        WHEN 't_ofbp' THEN '[tdmisg131200].[t_ofbp]'
                        WHEN 't_ofbp DESC' THEN '[tdmisg131200].[t_ofbp] DESC'
                        WHEN 't_vadr' THEN '[tdmisg131200].[t_vadr]'
                        WHEN 't_vadr DESC' THEN '[tdmisg131200].[t_vadr] DESC'
                        WHEN 't_padr' THEN '[tdmisg131200].[t_padr]'
                        WHEN 't_padr DESC' THEN '[tdmisg131200].[t_padr] DESC'
                        WHEN 't_dprj' THEN '[tdmisg131200].[t_dprj]'
                        WHEN 't_dprj DESC' THEN '[tdmisg131200].[t_dprj] DESC'
                        WHEN 't_user' THEN '[tdmisg131200].[t_user]'
                        WHEN 't_user DESC' THEN '[tdmisg131200].[t_user] DESC'
                        WHEN 't_date' THEN '[tdmisg131200].[t_date]'
                        WHEN 't_date DESC' THEN '[tdmisg131200].[t_date] DESC'
                        WHEN 't_subt' THEN '[tdmisg131200].[t_subt]'
                        WHEN 't_subt DESC' THEN '[tdmisg131200].[t_subt] DESC'
                        WHEN 't_refr' THEN '[tdmisg131200].[t_refr]'
                        WHEN 't_refr DESC' THEN '[tdmisg131200].[t_refr] DESC'
                        WHEN 't_appr' THEN '[tdmisg131200].[t_appr]'
                        WHEN 't_appr DESC' THEN '[tdmisg131200].[t_appr] DESC'
                        WHEN 't_rejc' THEN '[tdmisg131200].[t_rejc]'
                        WHEN 't_rejc DESC' THEN '[tdmisg131200].[t_rejc] DESC'
                        WHEN 't_rekm' THEN '[tdmisg131200].[t_rekm]'
                        WHEN 't_rekm DESC' THEN '[tdmisg131200].[t_rekm] DESC'
                        WHEN 't_apdt' THEN '[tdmisg131200].[t_apdt]'
                        WHEN 't_apdt DESC' THEN '[tdmisg131200].[t_apdt] DESC'
                        WHEN 't_apsu' THEN '[tdmisg131200].[t_apsu]'
                        WHEN 't_apsu DESC' THEN '[tdmisg131200].[t_apsu] DESC'
                        WHEN 't_irmk' THEN '[tdmisg131200].[t_irmk]'
                        WHEN 't_irmk DESC' THEN '[tdmisg131200].[t_irmk] DESC'
                        WHEN 't_iisu' THEN '[tdmisg131200].[t_iisu]'
                        WHEN 't_iisu DESC' THEN '[tdmisg131200].[t_iisu] DESC'
                        WHEN 't_retn' THEN '[tdmisg131200].[t_retn]'
                        WHEN 't_retn DESC' THEN '[tdmisg131200].[t_retn] DESC'
                        WHEN 't_smdt' THEN '[tdmisg131200].[t_smdt]'
                        WHEN 't_smdt DESC' THEN '[tdmisg131200].[t_smdt] DESC'
                        WHEN 't_isby' THEN '[tdmisg131200].[t_isby]'
                        WHEN 't_isby DESC' THEN '[tdmisg131200].[t_isby] DESC'
                        WHEN 't_isdt' THEN '[tdmisg131200].[t_isdt]'
                        WHEN 't_isdt DESC' THEN '[tdmisg131200].[t_isdt] DESC'
                        WHEN 't_Refcntd' THEN '[tdmisg131200].[t_Refcntd]'
                        WHEN 't_Refcntd DESC' THEN '[tdmisg131200].[t_Refcntd] DESC'
                        WHEN 't_Refcntu' THEN '[tdmisg131200].[t_Refcntu]'
                        WHEN 't_Refcntu DESC' THEN '[tdmisg131200].[t_Refcntu] DESC'
                        WHEN 't_edif' THEN '[tdmisg131200].[t_edif]'
                        WHEN 't_edif DESC' THEN '[tdmisg131200].[t_edif] DESC'
                        ELSE '[tdmisg131200].[t_tran]'
                    END
  EXEC (@LGSQL)

  SET @RecordCount = @@RowCount

  SELECT
    [tdmisg131200].*  
  FROM [tdmisg131200] 
      INNER JOIN #PageIndex
          ON [tdmisg131200].[t_tran] = #PageIndex.t_tran
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHInsert]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHInsert]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlHInsert]
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
  @Return_t_tran VarChar(9) = null OUTPUT 
  AS
  INSERT [tdmisg131200]
  (
   [t_tran]
  ,[t_type]
  ,[t_bpid]
  ,[t_cadr]
  ,[t_cprj]
  ,[t_logn]
  ,[t_subj]
  ,[t_remk]
  ,[t_issu]
  ,[t_stat]
  ,[t_ofbp]
  ,[t_vadr]
  ,[t_padr]
  ,[t_dprj]
  ,[t_user]
  ,[t_date]
  ,[t_subt]
  ,[t_refr]
  ,[t_appr]
  ,[t_rejc]
  ,[t_rekm]
  ,[t_apdt]
  ,[t_apsu]
  ,[t_irmk]
  ,[t_iisu]
  ,[t_retn]
  ,[t_smdt]
  ,[t_isby]
  ,[t_isdt]
  ,[t_Refcntd]
  ,[t_Refcntu]
  ,[t_edif]
  )
  VALUES
  (
   UPPER(@t_tran)
  ,@t_type
  ,@t_bpid
  ,@t_cadr
  ,@t_cprj
  ,@t_logn
  ,@t_subj
  ,@t_remk
  ,@t_issu
  ,@t_stat
  ,@t_ofbp
  ,@t_vadr
  ,@t_padr
  ,@t_dprj
  ,@t_user
  ,@t_date
  ,@t_subt
  ,@t_refr
  ,@t_appr
  ,@t_rejc
  ,@t_rekm
  ,@t_apdt
  ,@t_apsu
  ,@t_irmk
  ,@t_iisu
  ,@t_retn
  ,@t_smdt
  ,@t_isby
  ,@t_isdt
  ,@t_Refcntd
  ,@t_Refcntu
  ,@t_edif
  )
  SET @Return_t_tran = @t_tran
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHUpdate]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHUpdate]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlHDelete]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlHDelete]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlHDelete]
  @Original_t_tran VarChar(9),
  @RowCount int = null OUTPUT
  AS
  DELETE [tdmisg131200]
  WHERE
  [tdmisg131200].[t_tran] = @Original_t_tran
  SET @RowCount = @@RowCount
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDSelectByID]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDSelectByID]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDSelectListSearch]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDSelectListSearch]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlDSelectListSearch]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @KeyWord VarChar(250),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
    DECLARE @KeyWord1 VarChar(260)
    SET @KeyWord1 = '%' + LOWER(@KeyWord) + '%'
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_tran VarChar(9) NOT NULL
 ,t_docn VarChar(32) NOT NULL
 ,t_revn VarChar(20) NOT NULL
  )
  INSERT INTO #PageIndex (t_tran, t_docn, t_revn)
  SELECT [tdmisg132200].[t_tran], [tdmisg132200].[t_docn], [tdmisg132200].[t_revn] FROM [tdmisg132200]
 WHERE  
   ( 
         LOWER(ISNULL([tdmisg132200].[t_tran],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_docn],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_revn],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_stid],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_pono], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_remk],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_recv],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_refr],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([tdmisg132200].[t_rekm],'')) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_lock], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_recc], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_revd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_issu], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_Refcntd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([tdmisg132200].[t_Refcntu], 0)) LIKE @KeyWord1
   ) 
  ORDER BY
     CASE @OrderBy WHEN 't_tran' THEN [tdmisg132200].[t_tran] END,
     CASE @OrderBy WHEN 't_tran DESC' THEN [tdmisg132200].[t_tran] END DESC,
     CASE @OrderBy WHEN 't_docn' THEN [tdmisg132200].[t_docn] END,
     CASE @OrderBy WHEN 't_docn DESC' THEN [tdmisg132200].[t_docn] END DESC,
     CASE @OrderBy WHEN 't_revn' THEN [tdmisg132200].[t_revn] END,
     CASE @OrderBy WHEN 't_revn DESC' THEN [tdmisg132200].[t_revn] END DESC,
     CASE @OrderBy WHEN 't_stid' THEN [tdmisg132200].[t_stid] END,
     CASE @OrderBy WHEN 't_stid DESC' THEN [tdmisg132200].[t_stid] END DESC,
     CASE @OrderBy WHEN 't_pono' THEN [tdmisg132200].[t_pono] END,
     CASE @OrderBy WHEN 't_pono DESC' THEN [tdmisg132200].[t_pono] END DESC,
     CASE @OrderBy WHEN 't_remk' THEN [tdmisg132200].[t_remk] END,
     CASE @OrderBy WHEN 't_remk DESC' THEN [tdmisg132200].[t_remk] END DESC,
     CASE @OrderBy WHEN 't_recv' THEN [tdmisg132200].[t_recv] END,
     CASE @OrderBy WHEN 't_recv DESC' THEN [tdmisg132200].[t_recv] END DESC,
     CASE @OrderBy WHEN 't_refr' THEN [tdmisg132200].[t_refr] END,
     CASE @OrderBy WHEN 't_refr DESC' THEN [tdmisg132200].[t_refr] END DESC,
     CASE @OrderBy WHEN 't_redt' THEN [tdmisg132200].[t_redt] END,
     CASE @OrderBy WHEN 't_redt DESC' THEN [tdmisg132200].[t_redt] END DESC,
     CASE @OrderBy WHEN 't_rekm' THEN [tdmisg132200].[t_rekm] END,
     CASE @OrderBy WHEN 't_rekm DESC' THEN [tdmisg132200].[t_rekm] END DESC,
     CASE @OrderBy WHEN 't_lock' THEN [tdmisg132200].[t_lock] END,
     CASE @OrderBy WHEN 't_lock DESC' THEN [tdmisg132200].[t_lock] END DESC,
     CASE @OrderBy WHEN 't_recc' THEN [tdmisg132200].[t_recc] END,
     CASE @OrderBy WHEN 't_recc DESC' THEN [tdmisg132200].[t_recc] END DESC,
     CASE @OrderBy WHEN 't_revd' THEN [tdmisg132200].[t_revd] END,
     CASE @OrderBy WHEN 't_revd DESC' THEN [tdmisg132200].[t_revd] END DESC,
     CASE @OrderBy WHEN 't_issu' THEN [tdmisg132200].[t_issu] END,
     CASE @OrderBy WHEN 't_issu DESC' THEN [tdmisg132200].[t_issu] END DESC,
     CASE @OrderBy WHEN 't_Refcntd' THEN [tdmisg132200].[t_Refcntd] END,
     CASE @OrderBy WHEN 't_Refcntd DESC' THEN [tdmisg132200].[t_Refcntd] END DESC,
     CASE @OrderBy WHEN 't_Refcntu' THEN [tdmisg132200].[t_Refcntu] END,
     CASE @OrderBy WHEN 't_Refcntu DESC' THEN [tdmisg132200].[t_Refcntu] END DESC 

    SET @RecordCount = @@RowCount

  SELECT
    [tdmisg132200].*  
  FROM [tdmisg132200] 
      INNER JOIN #PageIndex
          ON [tdmisg132200].[t_tran] = #PageIndex.t_tran
          AND [tdmisg132200].[t_docn] = #PageIndex.t_docn
          AND [tdmisg132200].[t_revn] = #PageIndex.t_revn
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDSelectListFilteres]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDSelectListFilteres]
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
    SET @LGSQL = @LGSQL + ' AND [tdmisg132200].[t_docn] = ''' + @Filter_t_docn + ''''
  IF (@Filter_t_revn > '') 
    SET @LGSQL = @LGSQL + ' AND [tdmisg132200].[t_revn] = ''' + @Filter_t_revn + ''''
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
          ON [tdmisg132200].[t_tran] = #PageIndex.t_tran
          AND [tdmisg132200].[t_docn] = #PageIndex.t_docn
          AND [tdmisg132200].[t_revn] = #PageIndex.t_revn
  WHERE
        #PageIndex.IndexID > @StartRowIndex
        AND #PageIndex.IndexID < (@StartRowIndex + @MaximumRows + 1)
  ORDER BY
    #PageIndex.IndexID
  END
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDInsert]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDInsert]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlDInsert]
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
  @Return_t_tran VarChar(9) = null OUTPUT, 
  @Return_t_docn VarChar(32) = null OUTPUT, 
  @Return_t_revn VarChar(20) = null OUTPUT 
  AS
  INSERT [tdmisg132200]
  (
   [t_tran]
  ,[t_docn]
  ,[t_revn]
  ,[t_stid]
  ,[t_pono]
  ,[t_remk]
  ,[t_recv]
  ,[t_refr]
  ,[t_redt]
  ,[t_rekm]
  ,[t_lock]
  ,[t_recc]
  ,[t_revd]
  ,[t_issu]
  ,[t_Refcntd]
  ,[t_Refcntu]
  )
  VALUES
  (
   UPPER(@t_tran)
  ,UPPER(@t_docn)
  ,UPPER(@t_revn)
  ,@t_stid
  ,@t_pono
  ,@t_remk
  ,@t_recv
  ,@t_refr
  ,@t_redt
  ,@t_rekm
  ,@t_lock
  ,@t_recc
  ,@t_revd
  ,@t_issu
  ,@t_Refcntd
  ,@t_Refcntu
  )
  SET @Return_t_tran = @t_tran
  SET @Return_t_docn = @t_docn
  SET @Return_t_revn = @t_revn
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDUpdate]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDUpdate]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediTmtlDDelete]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediTmtlDDelete]
GO
 
CREATE PROCEDURE [dbo].[spediTmtlDDelete]
  @Original_t_tran VarChar(9),
  @Original_t_docn VarChar(32),
  @Original_t_revn VarChar(20),
  @RowCount int = null OUTPUT
  AS
  DELETE [tdmisg132200]
  WHERE
  [tdmisg132200].[t_tran] = @Original_t_tran
  AND [tdmisg132200].[t_docn] = @Original_t_docn
  AND [tdmisg132200].[t_revn] = @Original_t_revn
  SET @RowCount = @@RowCount
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibSelectByID]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibSelectByID]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibSelectListSearch]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibSelectListSearch]
GO
 
CREATE PROCEDURE [dbo].[spediALibSelectListSearch]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @KeyWord VarChar(250),
  @t_acti VarChar(1),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
    DECLARE @KeyWord1 VarChar(260)
    SET @KeyWord1 = '%' + LOWER(@KeyWord) + '%'
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_lbcd VarChar(50) NOT NULL
  )
  INSERT INTO #PageIndex (t_lbcd)
  SELECT [ttcisg127200].[t_lbcd] FROM [ttcisg127200]
 WHERE  
      [ttcisg127200].[t_acti] = (@t_acti)
   AND ( 
         LOWER(ISNULL([ttcisg127200].[t_lbcd],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg127200].[t_desc],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg127200].[t_serv],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg127200].[t_path],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg127200].[t_acti],'')) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg127200].[t_Refcntd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg127200].[t_Refcntu], 0)) LIKE @KeyWord1
   ) 
  ORDER BY
     CASE @OrderBy WHEN 't_lbcd' THEN [ttcisg127200].[t_lbcd] END,
     CASE @OrderBy WHEN 't_lbcd DESC' THEN [ttcisg127200].[t_lbcd] END DESC,
     CASE @OrderBy WHEN 't_desc' THEN [ttcisg127200].[t_desc] END,
     CASE @OrderBy WHEN 't_desc DESC' THEN [ttcisg127200].[t_desc] END DESC,
     CASE @OrderBy WHEN 't_serv' THEN [ttcisg127200].[t_serv] END,
     CASE @OrderBy WHEN 't_serv DESC' THEN [ttcisg127200].[t_serv] END DESC,
     CASE @OrderBy WHEN 't_path' THEN [ttcisg127200].[t_path] END,
     CASE @OrderBy WHEN 't_path DESC' THEN [ttcisg127200].[t_path] END DESC,
     CASE @OrderBy WHEN 't_acti' THEN [ttcisg127200].[t_acti] END,
     CASE @OrderBy WHEN 't_acti DESC' THEN [ttcisg127200].[t_acti] END DESC,
     CASE @OrderBy WHEN 't_Refcntd' THEN [ttcisg127200].[t_Refcntd] END,
     CASE @OrderBy WHEN 't_Refcntd DESC' THEN [ttcisg127200].[t_Refcntd] END DESC,
     CASE @OrderBy WHEN 't_Refcntu' THEN [ttcisg127200].[t_Refcntu] END,
     CASE @OrderBy WHEN 't_Refcntu DESC' THEN [ttcisg127200].[t_Refcntu] END DESC 

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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibSelectListFilteres]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibSelectListFilteres]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibInsert]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibInsert]
GO
 
CREATE PROCEDURE [dbo].[spediALibInsert]
  @t_lbcd VarChar(50),
  @t_desc VarChar(100),
  @t_serv VarChar(50),
  @t_path VarChar(50),
  @t_acti VarChar(1),
  @t_Refcntd Int,
  @t_Refcntu Int,
  @Return_t_lbcd VarChar(50) = null OUTPUT 
  AS
  INSERT [ttcisg127200]
  (
   [t_lbcd]
  ,[t_desc]
  ,[t_serv]
  ,[t_path]
  ,[t_acti]
  ,[t_Refcntd]
  ,[t_Refcntu]
  )
  VALUES
  (
   UPPER(@t_lbcd)
  ,@t_desc
  ,@t_serv
  ,@t_path
  ,@t_acti
  ,@t_Refcntd
  ,@t_Refcntu
  )
  SET @Return_t_lbcd = @t_lbcd
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibUpdate]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibUpdate]
GO
 
CREATE PROCEDURE [dbo].[spediALibUpdate]
  @Original_t_lbcd VarChar(50), 
  @t_lbcd VarChar(50),
  @t_desc VarChar(100),
  @t_serv VarChar(50),
  @t_path VarChar(50),
  @t_acti VarChar(1),
  @t_Refcntd Int,
  @t_Refcntu Int,
  @RowCount int = null OUTPUT
  AS
  UPDATE [ttcisg127200] SET 
   [t_lbcd] = @t_lbcd
  ,[t_desc] = @t_desc
  ,[t_serv] = @t_serv
  ,[t_path] = @t_path
  ,[t_acti] = @t_acti
  ,[t_Refcntd] = @t_Refcntd
  ,[t_Refcntu] = @t_Refcntu
  WHERE
  [t_lbcd] = @Original_t_lbcd
  SET @RowCount = @@RowCount
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediALibDelete]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediALibDelete]
GO
 
CREATE PROCEDURE [dbo].[spediALibDelete]
  @Original_t_lbcd VarChar(50),
  @RowCount int = null OUTPUT
  AS
  DELETE [ttcisg127200]
  WHERE
  [ttcisg127200].[t_lbcd] = @Original_t_lbcd
  SET @RowCount = @@RowCount
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesSelectByID]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesSelectByID]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesSelectListSearch]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesSelectListSearch]
GO
 
CREATE PROCEDURE [dbo].[spediASeriesSelectListSearch]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @KeyWord VarChar(250),
  @t_acti VarChar(1),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
    DECLARE @KeyWord1 VarChar(260)
    SET @KeyWord1 = '%' + LOWER(@KeyWord) + '%'
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_rnum Int NOT NULL
  )
  INSERT INTO #PageIndex (t_rnum)
  SELECT [ttcisg131200].[t_rnum] FROM [ttcisg131200]
 WHERE  
      [ttcisg131200].[t_acti] = (@t_acti)
   AND ( 
         LOWER(ISNULL([ttcisg131200].[t_seri],'')) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg131200].[t_rnum], 0)) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg131200].[t_acti],'')) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg131200].[t_Refcntd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg131200].[t_Refcntu], 0)) LIKE @KeyWord1
   ) 
  ORDER BY
     CASE @OrderBy WHEN 't_seri' THEN [ttcisg131200].[t_seri] END,
     CASE @OrderBy WHEN 't_seri DESC' THEN [ttcisg131200].[t_seri] END DESC,
     CASE @OrderBy WHEN 't_rnum' THEN [ttcisg131200].[t_rnum] END,
     CASE @OrderBy WHEN 't_rnum DESC' THEN [ttcisg131200].[t_rnum] END DESC,
     CASE @OrderBy WHEN 't_acti' THEN [ttcisg131200].[t_acti] END,
     CASE @OrderBy WHEN 't_acti DESC' THEN [ttcisg131200].[t_acti] END DESC,
     CASE @OrderBy WHEN 't_Refcntd' THEN [ttcisg131200].[t_Refcntd] END,
     CASE @OrderBy WHEN 't_Refcntd DESC' THEN [ttcisg131200].[t_Refcntd] END DESC,
     CASE @OrderBy WHEN 't_Refcntu' THEN [ttcisg131200].[t_Refcntu] END,
     CASE @OrderBy WHEN 't_Refcntu DESC' THEN [ttcisg131200].[t_Refcntu] END DESC 

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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesSelectListFilteres]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesSelectListFilteres]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesInsert]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesInsert]
GO
 
CREATE PROCEDURE [dbo].[spediASeriesInsert]
  @t_seri VarChar(50),
  @t_rnum Int,
  @t_acti VarChar(1),
  @t_Refcntd Int,
  @t_Refcntu Int,
  @Return_t_rnum Int = null OUTPUT 
  AS
  INSERT [ttcisg131200]
  (
   [t_seri]
  ,[t_rnum]
  ,[t_acti]
  ,[t_Refcntd]
  ,[t_Refcntu]
  )
  VALUES
  (
   @t_seri
  ,@t_rnum
  ,@t_acti
  ,@t_Refcntd
  ,@t_Refcntu
  )
  SET @Return_t_rnum = @t_rnum
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesUpdate]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesUpdate]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediASeriesDelete]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediASeriesDelete]
GO
 
CREATE PROCEDURE [dbo].[spediASeriesDelete]
  @Original_t_rnum Int,
  @RowCount int = null OUTPUT
  AS
  DELETE [ttcisg131200]
  WHERE
  [ttcisg131200].[t_rnum] = @Original_t_rnum
  SET @RowCount = @@RowCount
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileSelectByID]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileSelectByID]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileSelectListSearch]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileSelectListSearch]
GO
 
CREATE PROCEDURE [dbo].[spediAFileSelectListSearch]
  @LoginID NVarChar(8),
  @StartRowIndex int,
  @MaximumRows int,
  @KeyWord VarChar(250),
  @OrderBy NVarChar(50),
  @RecordCount Int = 0 OUTPUT
  AS
  BEGIN
    DECLARE @KeyWord1 VarChar(260)
    SET @KeyWord1 = '%' + LOWER(@KeyWord) + '%'
  CREATE TABLE #PageIndex (
  IndexID INT IDENTITY (1, 1) NOT NULL
 ,t_drid VarChar(50) NOT NULL
  )
  INSERT INTO #PageIndex (t_drid)
  SELECT [ttcisg132200].[t_drid] FROM [ttcisg132200]
 WHERE  
   ( 
         LOWER(ISNULL([ttcisg132200].[t_drid],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_dcid],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_hndl],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_indx],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_prcd],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_fnam],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_lbcd],'')) LIKE @KeyWord1
     OR LOWER(ISNULL([ttcisg132200].[t_atby],'')) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg132200].[t_Refcntd], 0)) LIKE @KeyWord1
     OR STR(ISNULL([ttcisg132200].[t_Refcntu], 0)) LIKE @KeyWord1
   ) 
  ORDER BY
     CASE @OrderBy WHEN 't_drid' THEN [ttcisg132200].[t_drid] END,
     CASE @OrderBy WHEN 't_drid DESC' THEN [ttcisg132200].[t_drid] END DESC,
     CASE @OrderBy WHEN 't_dcid' THEN [ttcisg132200].[t_dcid] END,
     CASE @OrderBy WHEN 't_dcid DESC' THEN [ttcisg132200].[t_dcid] END DESC,
     CASE @OrderBy WHEN 't_hndl' THEN [ttcisg132200].[t_hndl] END,
     CASE @OrderBy WHEN 't_hndl DESC' THEN [ttcisg132200].[t_hndl] END DESC,
     CASE @OrderBy WHEN 't_indx' THEN [ttcisg132200].[t_indx] END,
     CASE @OrderBy WHEN 't_indx DESC' THEN [ttcisg132200].[t_indx] END DESC,
     CASE @OrderBy WHEN 't_prcd' THEN [ttcisg132200].[t_prcd] END,
     CASE @OrderBy WHEN 't_prcd DESC' THEN [ttcisg132200].[t_prcd] END DESC,
     CASE @OrderBy WHEN 't_fnam' THEN [ttcisg132200].[t_fnam] END,
     CASE @OrderBy WHEN 't_fnam DESC' THEN [ttcisg132200].[t_fnam] END DESC,
     CASE @OrderBy WHEN 't_lbcd' THEN [ttcisg132200].[t_lbcd] END,
     CASE @OrderBy WHEN 't_lbcd DESC' THEN [ttcisg132200].[t_lbcd] END DESC,
     CASE @OrderBy WHEN 't_atby' THEN [ttcisg132200].[t_atby] END,
     CASE @OrderBy WHEN 't_atby DESC' THEN [ttcisg132200].[t_atby] END DESC,
     CASE @OrderBy WHEN 't_aton' THEN [ttcisg132200].[t_aton] END,
     CASE @OrderBy WHEN 't_aton DESC' THEN [ttcisg132200].[t_aton] END DESC,
     CASE @OrderBy WHEN 't_Refcntd' THEN [ttcisg132200].[t_Refcntd] END,
     CASE @OrderBy WHEN 't_Refcntd DESC' THEN [ttcisg132200].[t_Refcntd] END DESC,
     CASE @OrderBy WHEN 't_Refcntu' THEN [ttcisg132200].[t_Refcntu] END,
     CASE @OrderBy WHEN 't_Refcntu DESC' THEN [ttcisg132200].[t_Refcntu] END DESC 

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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileSelectListFilteres]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileSelectListFilteres]
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileInsert]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileInsert]
GO
 
CREATE PROCEDURE [dbo].[spediAFileInsert]
  @t_drid VarChar(50),
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
   UPPER(@t_drid)
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
  SET @Return_t_drid = @t_drid
  GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileUpdate]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileUpdate]
GO
 
CREATE PROCEDURE [dbo].[spediAFileUpdate]
  @Original_t_drid VarChar(50), 
  @t_drid VarChar(50),
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
   [t_drid] = @t_drid
  ,[t_dcid] = @t_dcid
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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spediAFileDelete]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spediAFileDelete]
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

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO
