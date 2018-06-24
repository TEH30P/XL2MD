param
(	[parameter(Mandatory=1, Position=0)][String]$iPathXLS
,	[parameter(Mandatory=0, Position=1)][String]$iPathMD)

[Text.StringBuilder]$gSBOut = New-Object Text.StringBuilder(1MB);
#--------------------------#
function XLFontReadInfo([psobject]$XLParent)
{	if ($null -eq $XLParent)
	{	return New-Object psobject -Property @{Strikethrough = $false; Italic=$false; Bold=$false; Underline = [Microsoft.Office.Interop.Excel.XlUnderlineStyle]::xlUnderlineStyleNone}}
	
	$XLFont = $XLParent.Font;
	
	New-Object psobject -Property @{Strikethrough = $XLFont.Strikethrough; Italic = $XLFont.Italic; Bold = $XLFont.Bold; Underline = $XLFont.Underline};
	
	Get-Variable XLFont -Scope Local | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
	return;
}
#--------------------------#
#[Text.StringBuilder]$gSBOut
function CellTextDecorate([psobject]$FontCurr, [psobject]$FontNew)
{	if ($FontCurr.Bold -ne $FontNew.Bold)
	{	[Void]$gSBOut.Append('**')}
	
	if ($FontCurr.Italic -ne $FontNew.Italic)
	{	[Void]$gSBOut.Append('*')}
	
	if ($FontCurr.Strikethrough -ne $FontNew.Strikethrough)
	{	[Void]$gSBOut.Append('~~')}
	
	if (([Microsoft.Office.Interop.Excel.XlUnderlineStyle]::xlUnderlineStyleNone -eq $FontCurr.Underline) -xor ([Microsoft.Office.Interop.Excel.XlUnderlineStyle]::xlUnderlineStyleNone -eq $FontCurr.Underline))
	{	[Void]$gSBOut.Append('__')}
}
#--------------------------#
#[Text.StringBuilder]$gSBOut
function CellCharTranslate([String]$iCh, [String]$iChPrev)
{	[Char]$Ch = $iCh;
	[String]$Ch2 = $iChPrev + $iCh;
	#[Void]$gSBOut.Clear();
	
	switch -exact ($Ch)
	{	"`n"
		{	[Void]$gSBOut.Append('<br>')}
		{$Ch2 -ceq "`r`n"}
		{	[Void]$gSBOut.Append('<br>')}
		{$Ch2.Length -eq 2 -and $Ch2.StartsWith("`r")}
		{	[Void]$gSBOut.Append('<br>').Append($Ch)}
		{'|*~_'.Contains($_)}
		{	[Void]$gSBOut.Append([Char]'\').Append($Ch)}
		default
		{	[Void]$gSBOut.Append($Ch)}
	}
}
#--------------------------#
############################
try
{	[String]$PathXLS = Convert-Path -LiteralPath $iPathXLS;
	
	if ([String]::IsNullOrEmpty($iPathMD))
	{	[String]$PathMDBase = [IO.Path]::Combine([IO.Path]::GetDirectoryName($iPathXLS), [IO.Path]::GetFileNameWithoutExtension($iPathXLS))}
	else
	{	[String]$PathMDBase = [IO.Path]::Combine([IO.Path]::GetDirectoryName($iPathMD), [IO.Path]::GetFileNameWithoutExtension($iPathMD))}
	
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel;
	
	try
	{	[psobject]$XL = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}
	catch
	{	}

	if ($null -ne $XL)
	{	[Boolean]$IsXLBackground = !$XL.Visible
		
		[psobject]$XLWBTrg = $null;
		
		$XL.Workbooks | ? {$_.FullName -eq $PathXLS} | % {$XLWBTrg = $_};
		
		if ($null -eq $XLWBTrg)
		{	rv XL}
	}
	
	if ($null -eq $XL)
	{	[Boolean]$IsXLBackground = $true;
		[Microsoft.Office.Interop.Excel.Application]$XL = New-Object Microsoft.Office.Interop.Excel.ApplicationClass;
		<#[Microsoft.Office.Interop.Excel.Workbook]#>$XLWBTrg = $XL.Workbooks.Open($PathXLS);
		$XL.Visible = $false;
	}
	
	[String]$ProgressActivity = 'Converting: "' + $PathXLS + '"';
	
	foreach ($XLWSIt in $XLWBTrg.Worksheets)
	{	[Void]$gSBOut.Clear();
		[String]$PathMD = "$PathMDBase.$($XLWSIt.Name).md";
		$XLRNData = $XLWSIt.UsedRange;
		[Int32]$RowCnt = $XLRNData.Rows.Count + 1;
		[Int32]$ColCnt = $XLRNData.Columns.Count;
		[Collections.ArrayList]$aBorderPosRow = New-Object Collections.ArrayList($ColCnt);
		
		[Int32[]]$aBorderPosCol = $null;
		
		for ([Int32]$ColIdx = 0; $ColIdx -lt $ColCnt; $ColIdx++)
		{	[Array]::Resize([ref]$aBorderPosCol, $RowCnt);
			[Void]$aBorderPosRow.Add($aBorderPosCol);
			$aBorderPosCol = $null;
		}
		
		[Int32[]]$aBorderPosCol0 = $null;
		[Array]::Resize([ref]$aBorderPosCol0, $RowCnt);
		
		
		[Int32]$ProgressMax = $XLRNData.Rows.Count;
		[Int32]$Progress = 0;
		[String]$ProgressXLWorksheet = $XLWSIt.Name;
		Write-Progress -Activity $ProgressActivity -Status "$Progress/$ProgressMax rows processed" -CurrentOperation "Processing worksheet: `"$($ProgressXLWorksheet)`"" -PercentComplete ([Math]::Round($Progress/$ProgressMax*100.));
		
		[Int32]$RowIdx = 0;
		[Boolean]$HeadRow = $true;
		
		foreach ($XLRNRow in $XLRNData.Rows)
		{	$Progress++;
			
			if ($XLRNRow.EntireRow.Hidden)
			{	continue}
			
			[Int32]$ColIdx = 0;
			
			[Void]$gSBOut.Append('|');
			$aBorderPosCol0[$RowIdx] = $gSBOut.Length - 1;
			
			foreach ($XLRNCell in $XLRNRow.Cells())
			{	if ($XLRNRow.EntireColumn.Hidden)
				{	continue}
				
				[Void]$gSBOut.Append(' ');
				[psobject]$CellFont = XLFontReadInfo;
				[psobject]$CellFontIt = XLFontReadInfo $XLRNCell;
				
				#if (($XLRNCell.HasFormula) -or ($null -eq $XLRNCell.Characters($null, $null).Count))
				if ($XLRNCell.HasFormula -or $null -eq $CellFontIt.Bold -or $null -eq $CellFontIt.Italic -or $null -eq $CellFontIt.Strikethrough -or $null -eq $CellFontIt.Underline)
				{	#[psobject]$CellFontIt = XLFontReadInfo $XLRNCell;
					CellTextDecorate $CellFont  $CellFontIt;
					
					([String]$XLRNCell.Value()).ToCharArray() `
					|	% -Begin {[String]$ChPrev = [String]::Empty} -Process {CellCharTranslate $_ $ChPrev; $ChPrev = $_};
					
					CellTextDecorate $CellFontIt $CellFont;
				}
				else
				{	[String]$ChPrev = [String]::Empty;
					
					for ([Int32]$k = 1; $k -le $XLRNCell.Characters($null, $null).Count; $k++)
					{	[psobject]$CellFontIt = XLFontReadInfo $XLRNCell.Characters($k, 1);
						CellTextDecorate $CellFont $CellFontIt;
						[String]$Ch = $XLRNCell.Characters($k, 1).Text;
						CellCharTranslate $Ch $ChPrev;
						$ChPrev = $Ch;
						$CellFont = $CellFontIt;
					}
					
					[psobject]$CellFontIt = $CellFont;
					[psobject]$CellFont = XLFontReadInfo;
					CellTextDecorate $CellFont $CellFontIt;
				}
				
				[Void]$gSBOut.Append(' |');
				$aBorderPosRow[$ColIdx][$RowIdx] = $gSBOut.Length - 1;
				$ColIdx++;
				
				Get-Variable XLRNCell | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
			}
			
			[Void]$gSBOut.AppendLine();
			
			if ($HeadRow)
			{	$HeadRow = $false;
				$RowIdx++;
				[Int32]$ColIdx = 0;
				
				[Void]$gSBOut.Append('|');
				$aBorderPosCol0[$RowIdx] = $gSBOut.Length - 1;
				
				foreach ($XLRNCell in $XLRNRow.Cells())
				{	switch -exact ($XLRNCell.HorizontalAlignment)
					{	[Microsoft.Office.Interop.Excel.Constants]::xlRight
						{	[Void]$gSBOut.Append(' ---: |')}
						[Microsoft.Office.Interop.Excel.Constants]::xlCenter
						{	[Void]$gSBOut.Append(' :---: |')}
						default
						{	[Void]$gSBOut.Append(' :--- |')}
					}
					
					$aBorderPosRow[$ColIdx][$RowIdx] = $gSBOut.Length - 1;
					$ColIdx++;
					
					Get-Variable XLRNCell | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
				}
				
				[Void]$gSBOut.AppendLine();
			}
		
			Get-Variable XLRNRow | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
			
			$ColCnt = $ColIdx;
			
			$RowIdx++;
			Write-Progress -Activity $ProgressActivity -Status "$Progress/$ProgressMax rows processed" -CurrentOperation "Processing worksheet: `"$($ProgressXLWorksheet)`"" -PercentComplete ([Math]::Round($Progress/$ProgressMax*100.));
		}
		
		$RowCnt = $RowIdx;
		
		Get-Variable XLWSIt | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
		
		if (-not ($RowCnt -and $ColCnt))
		{	Write-Warning "The worksheet `"$($ProgressXLWorksheet)`" have no visible rows and/or columns."
			continue;
		}
		#!!!REM: Does not work. Debug needed.
		
		for([Int32]$ColIdx = 0; $ColIdx -lt $ColCnt; $ColIdx++)
		{	[Int32[]]$aBorderPosCol = $aBorderPosRow[$ColIdx];
			[Int32]$CellMax = 0;
			
			for([Int32]$RowIdx = 0; $RowIdx -lt $RowCnt; $RowIdx++)
			{	if ($CellMax -lt $aBorderPosCol[$RowIdx] - $aBorderPosCol0[$RowIdx])
				{	$CellMax = $aBorderPosCol[$RowIdx] - $aBorderPosCol0[$RowIdx]}
			}
			
			for([Int32]$RowIdx = 0; $RowIdx -lt $RowCnt; $RowIdx++)
			{	[Int32]$Offset = $CellMax - ($aBorderPosCol[$RowIdx] - $aBorderPosCol0[$RowIdx]);
				
				if ($Offset)
				{	if ($RowIdx -eq 1)
						# Line like "| --- | ... | --- |"
					{	[Void]$gSBOut.Insert($aBorderPosCol[$RowIdx] - 3, '-', $Offset)}
					else
					{	[Void]$gSBOut.Insert($aBorderPosCol[$RowIdx], ' ', $Offset)};
					
					for ([Int32]$RowIdx1 = $RowIdx + 1; $RowIdx1 -lt $RowCnt; $RowIdx1++)
					{	$aBorderPosCol0[$RowIdx1] += $Offset}
						
					
					for ([Int32]$ColIdx1 = 0; $ColIdx1 -lt $ColIdx; $ColIdx1++)
					{	[Int32[]]$aBorderPosCol1 = $aBorderPosRow[$ColIdx1];
						
						for ([Int32]$RowIdx1 = $RowIdx + 1; $RowIdx1 -lt $RowCnt; $RowIdx1++)
						{	$aBorderPosCol1[$RowIdx1] += $Offset}
					}
					
					for ([Int32]$ColIdx1 = $ColIdx; $ColIdx1 -lt $ColCnt; $ColIdx1++)
					{	[Int32[]]$aBorderPosCol1 = $aBorderPosRow[$ColIdx1];
						
						for ([Int32]$RowIdx1 = $RowIdx; $RowIdx1 -lt $RowCnt; $RowIdx1++)
						{	$aBorderPosCol1[$RowIdx1] += $Offset}
					}
				}
			}
		}
		#>
		
		Write-Progress -Activity $ProgressActivity -Status "$ProgressMax/$ProgressMax rows processed" -CurrentOperation "Saving MD worksheet: `"$($ProgressXLWorksheet)`"" -PercentComplete 100;
		
		#Write-Host "Worksheet `"$($ProgressXLWorksheet)`" MD size = $($gSBOut.Length)";
		
		[IO.StreamWriter]$StrWr = [IO.File]::CreateText($PathMD);
		[Char[]]$aChBuff = $null;
		[Array]::Resize([Ref]$aChBuff, 64kb)
		[Int32]$Offset = 0;
		[Int32]$Len = [Math]::Min($gSBOut.Length, 64kb);
		
		while ($Len)
		{	$gSBOut.CopyTo($Offset, $aChBuff, 0, $Len);
			$StrWr.Write($aChBuff, 0, $Len);	
			$Offset += $Len;
			[Int32]$Len = [Math]::Min($gSBOut.Length - $Offset, 64kb);
		}
		
		$StrWr.Close(); $StrWr = $null;
		
		#!!!DBG: break at first page.
		#break;
	}	
}
catch
{	throw}
finally
{	if ($null -ne $XL)
	{	if ($IsXLBackground)
		{	$XL.Workbooks | % {$_.Close($false)}
			$XL.Quit()
		}
		
		Get-Variable XL | % {[Void][Runtime.InteropServices.Marshal]::ReleaseComObject($_.Value); $_} | Remove-Variable;
	}
	
	if ($StrWr -ne $null)
	{	$StrWr.Close()}
}
