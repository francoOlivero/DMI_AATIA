DROP TABLE IF EXISTS #ATM_Use
DECLARE @targetModule NVARCHAR(100);

SET @targetModule = 'Regular Life'
	
SELECT *
INTO #ATM_Use
FROM (
	SELECT DISTINCT *,
	ISNULL(Usage,'Product Feature') as Table_Usage_Type
	FROM Obra.dbo.Assumption_Tables_Where_Used
	WHERE	AsUsage in (0,1,5) 
		and [Used By] not in ('Cell Override', 'Reinsurance Terms')
		and [Module] = @targetModule
) AS SourceQuery;


Select * from #ATM_Use
select * from obra.dbo.Assumption_Inventory_Tables

SELECT 
	t1.[Section],
	t1.[Shape],
	t1.[TableName],
	t1.[Row],
	t1.[Op],
	t1.[LnkSection],
	t1.[LnkTable],
	t1.[FORMULA],
	t2.[Name],
	t2.[Obj Name], 
	t2.[Used By],
	t2.[Module], 
	t2.[Table Type], 
	t2.[Table_Usage_Type],
	t1.[C1], t1.[C2], t1.[C3], t1.[C4], t1.[C5], t1.[C6], t1.[C7], t1.[C8], t1.[C9], t1.[C10],
	t1.[C11], t1.[C12], t1.[C13], t1.[C14], t1.[C15], t1.[C16], t1.[C17], t1.[C18], t1.[C19], t1.[C20],
	t1.[C21], t1.[C22], t1.[C23], t1.[C24], t1.[C25], t1.[C26], t1.[C27], t1.[C28], t1.[C29], t1.[C30],
	t1.[C31], t1.[C32], t1.[C33], t1.[C34], t1.[C35], t1.[C36], t1.[C37], t1.[C38], t1.[C39], t1.[C40],
	t1.[C41], t1.[C42], t1.[C43], t1.[C44], t1.[C45], t1.[C46], t1.[C47], t1.[C48], t1.[C49], t1.[C50],
	t1.[C51], t1.[C52], t1.[C53], t1.[C54], t1.[C55], t1.[C56], t1.[C57], t1.[C58], t1.[C59], t1.[C60],
	t1.[C61], t1.[C62], t1.[C63], t1.[C64], t1.[C65], t1.[C66], t1.[C67], t1.[C68], t1.[C69], t1.[C70],
	t1.[C71], t1.[C72], t1.[C73], t1.[C74], t1.[C75], t1.[C76], t1.[C77], t1.[C78], t1.[C79], t1.[C80],
	t1.[C81], t1.[C82], t1.[C83], t1.[C84], t1.[C85], t1.[C86], t1.[C87], t1.[C88], t1.[C89], t1.[C90],
	t1.[C91], t1.[C92], t1.[C93], t1.[C94], t1.[C95], t1.[C96], t1.[C97], t1.[C98], t1.[C99], t1.[C100],
	t1.[C101], t1.[C102], t1.[C103], t1.[C104], t1.[C105], t1.[C106], t1.[C107], t1.[C108], t1.[C109], t1.[C110],
	t1.[C111], t1.[C112], t1.[C113], t1.[C114], t1.[C115], t1.[C116], t1.[C117], t1.[C118], t1.[C119], t1.[C120], t1.[C121]

FROM Obra.dbo.Assumption_Inventory_Tables AS t1
RIGHT JOIN #ATM_Use AS t2 --Keep only tables that are being used in the model
	ON t1.[TableName] = t2.[Name];
