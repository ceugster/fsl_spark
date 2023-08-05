package ch.eugster.filemaker.fsl.xls;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public class CellRange
{
	private Sheet sheet;
	
	private CellRangeAddress range;
	
	public CellRange(Sheet sheet, CellRangeAddress range)
	{
		this.sheet = sheet;
		this.range = range;
	}
	
	public CellRange(Sheet sheet, CellAddress topLeft, CellAddress bottomRight)
	{
		this.sheet = sheet;
		this.range = new CellRangeAddress(topLeft.getRow(), topLeft.getColumn(), bottomRight.getRow(), bottomRight.getColumn());
	}
	
	public CellRange(Sheet sheet, int top, int left, int bottom, int right)
	{
		this.sheet = sheet;
		this.range = new CellRangeAddress(top, left, bottom, right);
	}
	
	public Sheet getSheet()
	{
		return this.sheet;
	}
	
	public CellRangeAddress getRange()
	{
		return this.range;
	}
}
