package ch.eugster.filemaker.fsl.xls;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import com.fasterxml.jackson.databind.node.ObjectNode;

enum Direction implements MessageProvider
{
	DOWN("down"), LEFT("left"), RIGHT("right"), UP("up"), DEFAULT("right");

	private Direction(String value)
	{
		this.value = value;
	}

	private String value;

	public String direction()
	{
		return this.value;
	}

	public CellAddress nextIndex(CellAddress cellAddress)
	{
		switch (this)
		{
			case LEFT:
				return new CellAddress(cellAddress.getRow(), cellAddress.getColumn() - 1);
			case RIGHT:
				return new CellAddress(cellAddress.getRow(), cellAddress.getColumn() + 1);
			case UP:
				return new CellAddress(cellAddress.getRow() - 1, cellAddress.getColumn());
			case DOWN:
				return new CellAddress(cellAddress.getRow() + 1, cellAddress.getColumn());
			case DEFAULT:
				return new CellAddress(cellAddress.getRow(), cellAddress.getColumn() + 1);
			default:
				return new CellAddress(cellAddress.getRow(), cellAddress.getColumn() + 1);
		}
	}

	public boolean validRange(ObjectNode responseNode, Workbook workbook, CellAddress cellAddress, int numberOfCells)
	{
		switch (this)
		{
			case LEFT:
			{
				boolean valid = cellAddress.getColumn() - numberOfCells + 1 < 0 ? false : true;
				if (!valid)
				{
					addErrorMessage(responseNode, "minimal_horizontal_cell_position exceeds sheet's extent negatively");
				}
				return valid;
			}
			case RIGHT:
			{
				boolean valid = cellAddress.getColumn() + numberOfCells - 1 > workbook
						.getSpreadsheetVersion().getLastColumnIndex() ? false : true;
				if (!valid)
				{
					addErrorMessage(responseNode, "maximal_horizontal_cell_position exceeds sheet's extent positively");
				}
				return valid;
			}
			case UP:
			{
				boolean valid = cellAddress.getRow() - numberOfCells + 1 < 0 ? false : true;
				if (!valid)
				{
					addErrorMessage(responseNode, "minimal_vertical_cell_position exceeds sheet's extent negatively");
				}
				return valid;
			}
			case DOWN:
			{
				boolean valid = cellAddress.getRow() + numberOfCells - 1 > workbook
						.getSpreadsheetVersion().getLastRowIndex() ? false : true;
				if (!valid)
				{
					addErrorMessage(responseNode, "maximal_vertical_cell_position exceeds sheet's extent positively");
				}
				return valid;
			}
			case DEFAULT:
			{
				return Direction.RIGHT.validRange(responseNode, workbook, cellAddress, numberOfCells);
			}
			default:
				return false;
		}
	}
}

