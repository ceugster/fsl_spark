package ch.eugster.filemaker.fsl.xls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.lang.reflect.Parameter;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.ptg.AreaPtgBase;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.IntNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.TextNode;

import ch.eugster.filemaker.fsl.Executor;

/**
 * @author christian
 *
 * @created 2023-07-24
 * 
 * @updated
 * 
 *          The public methods have to follow this convention:
 * 
 *          - There is always one parameter of type string. This string must be
 *          a valid json object, that is a conversion to a jackson object node
 *          must be successfully done. The json object can contain zero or more
 *          attributes of valid json types, depending on the method called (see
 *          method descriptions). - There is always on return parameter of type
 *          string, This string too must be a valid json object as above. The
 *          json attribute 'status' is mandatory and contains either 'OK' or
 *          'Fehler', depending on the result of the method. Occuring errors are
 *          documented in an array object named 'errors'. Depending on the
 *          method json attributes with information are returned. Valid
 *          attribute names are documented at the method.
 * 
 *          There is a set of controlled attributes, that are recognized
 *          valid. @see ch.eugster.filemaker.fsl.xls.Key
 */
public class Xls extends Executor
{
	public static Workbook activeWorkbook;

	/**
	 * Set active sheet
	 * 
	 * @param sheet name (string) or index (integer)
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void activateSheet()
	{
		if (isWorkbookPresent())
		{
			doActivateSheet();
		}
	}

	/**
	 * Return active sheet name and index
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void activeSheet()
	{
		if (isWorkbookPresent())
		{
			doGetActiveSheet();
		}
	}

	/**
	 * Return active sheet name and index
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void activeSheetPresent()
	{
		if (isWorkbookPresent())
		{
			doActiveSheetPresent();
		}
	}

	public void applyCellStyles()
	{
		if (isWorkbookPresent())
		{
			doApplyCellStyles();
		}
	}

	public void applyFontStyles()
	{
		if (isWorkbookPresent())
		{
			doApplyFontStyles();
		}
	}

	public void autoSizeColumns()
	{
		if (isWorkbookPresent())
		{
			doAutoSizeColumns();
		}
	}

	public void copyCells()
	{
		if (isWorkbookPresent())
		{
			doCopyCells();
		}
	}

	/**
	 * Create sheet
	 * 
	 * @param sheet name (optional)
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void createSheet()
	{
		if (isWorkbookPresent())
		{
			doCreateSheet();
		}
	}

	/**
	 * Creates a workbook
	 * 
	 * @param workbook name
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void createWorkbook()
	{
		doCreateWorkbook();
	}

	/**
	 * Create a workbook with initial sheet
	 * 
	 * @param sheet name (optional)
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void createWorkbookWithSheet()
	{
		if (doCreateWorkbook())
		{
			doCreateSheet();
		}
	}

	/**
	 * Drop sheet
	 * 
	 * @param sheet name (string) or index (integer)
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void dropSheet()
	{
		if (isWorkbookPresent())
		{
			doDropSheet();
		}
	}

	/**
	 * Returns an string array of all callable methods
	 * 
	 * @param getRequestNode() empty
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return methods string array of method names
	 * @return optional errors containing error messages
	 * 
	 */
	public void getCallableMethods()
	{
		ArrayNode callableMethods = getResponseNode().arrayNode();
		Method[] methods = Xls.class.getDeclaredMethods();
		for (Method method : methods)
		{
			if (Modifier.isPublic(method.getModifiers()) && Modifier.isStatic(method.getModifiers()))
			{
				Parameter[] parameters = method.getParameters();
				if (parameters.length == 1 && parameters[0].getType().equals(String.class))
				{
					callableMethods.add(method.getName());
				}
			}
		}
		getResponseNode().set(Executor.RESULT, callableMethods);
	}

	public void getSupportedFunctionNames()
	{
		ArrayNode arrayNode = getResponseNode().arrayNode();
		Collection<String> supportedFunctionNames = FunctionEval.getSupportedFunctionNames();
		for (String supportedFunctionName : supportedFunctionNames)
		{
			arrayNode.add(supportedFunctionName);
		}
		getResponseNode().set(Executor.RESULT, arrayNode);
	}

	/**
	 * Rename existing sheet
	 * 
	 * @param index (integer)
	 * @param sheet (string) new sheet name
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void moveSheet()
	{
		if (isWorkbookPresent())
		{
			doMoveSheet();
		}
	}

	/**
	 * Release current workbook
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void releaseWorkbook()
	{
		if (isWorkbookPresent())
		{
			doReleaseWorkbook();
		}
	}

	/**
	 * Rename existing sheet
	 * 
	 * @param index (integer)
	 * @param sheet (string) new sheet name
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void renameSheet()
	{
		if (isWorkbookPresent())
		{
			doRenameSheet();
		}
	}

	public void rotateCells()
	{
		if (isWorkbookPresent())
		{
			doRotateCells();
		}
	}

	/**
	 * Save and release current workbook
	 * 
	 * @param path where to save the workbook
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void saveAndReleaseWorkbook()
	{
		if (isWorkbookPresent())
		{
			if (doSaveWorkbook())
			{
				doReleaseWorkbook();
			}
		}
	}

	/**
	 * Save current workbook
	 * 
	 * @param path where to save the workbook
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void saveWorkbook()
	{
		if (isWorkbookPresent())
		{
			doSaveWorkbook();
		}
	}

	public void setCells()
	{
		if (isWorkbookPresent())
		{
			doSetCells();
		}
	}

	public void setFooters()
	{
		if (isWorkbookPresent())
		{
			doSetFooters();
		}
	}

	public void setHeaders()
	{
		if (isWorkbookPresent())
		{
			doSetHeaders();
		}
	}

	public void setPrintSetup()
	{
		if (isWorkbookPresent())
		{
			doSetPrintSetup();
		}
	}

	/**
	 * Return list of sheet names by index order
	 * 
	 * @return status 'OK' or 'Fehler'
	 * @return optional 'errors' array of error messages
	 * 
	 */
	public void sheetNames()
	{
		if (isWorkbookPresent())
		{
			doGetSheetNames();
		}
	}

	public void workbookPresent()
	{
		getResponseNode().put(Executor.RESULT, Objects.nonNull(Xls.activeWorkbook) ? 1 : 0);
	}

	private void copyCell(Cell sourceCell, Cell targetCell)
	{
		CellType cellType = sourceCell.getCellType();
		switch (cellType)
		{
			case BLANK:
				break;
			case _NONE:
				break;
			case FORMULA:
			{
				String formula = sourceCell.getCellFormula();
				CellAddress sourceCellAddress = new CellAddress(sourceCell);
				CellAddress targetCellAddress = new CellAddress(targetCell);
				int rowDiff = targetCellAddress.getRow() - sourceCellAddress.getRow();
				int colDiff = targetCellAddress.getColumn() - sourceCellAddress.getColumn();
				formula = copyFormula(sourceCell.getRow().getSheet(), formula, rowDiff, colDiff);
				targetCell.setCellFormula(formula);
				break;
			}
			default:
			{
				CellUtil.copyCell(sourceCell, targetCell, null, null);
				break;
			}
		}
	}

	private String copyFormula(Sheet sheet, String formula, int rowDiff, int colDiff)
	{
		FormulaParsingWorkbook workbookWrapper = getFormulaParsingWorkbook(sheet);
		Ptg[] ptgs = FormulaParser.parse(formula, workbookWrapper, FormulaType.CELL,
				sheet.getWorkbook().getSheetIndex(sheet));
		for (int i = 0; i < ptgs.length; i++)
		{
			if (ptgs[i] instanceof RefPtgBase)
			{
				// base class for cell references
				RefPtgBase ref = (RefPtgBase) ptgs[i];
				if (ref.isRowRelative())
				{
					ref.setRow(ref.getRow() + rowDiff);
				}
				if (ref.isColRelative())
				{
					ref.setColumn(ref.getColumn() + colDiff);
				}
			}
			else if (ptgs[i] instanceof AreaPtgBase)
			{
				// base class for range references
				AreaPtgBase ref = (AreaPtgBase) ptgs[i];
				if (ref.isFirstColRelative())
				{
					ref.setFirstColumn(ref.getFirstColumn() + colDiff);
				}
				if (ref.isLastColRelative())
				{
					ref.setLastColumn(ref.getLastColumn() + colDiff);
				}
				if (ref.isFirstRowRelative())
				{
					ref.setFirstRow(ref.getFirstRow() + rowDiff);
				}
				if (ref.isLastRowRelative())
				{
					ref.setLastRow(ref.getLastRow() + rowDiff);
				}
			}
		}

		formula = FormulaRenderer.toFormulaString(getFormulaRenderingWorkbook(sheet), ptgs);
		return formula;
	}

	private boolean doActivateSheet()
	{
		boolean result = true;
		JsonNode sheetNode = getRequestNode().findPath(Key.SHEET.key());
		if (sheetNode.isTextual())
		{
			Sheet sheet = activeWorkbook.getSheet(sheetNode.asText());
			if (Objects.nonNull(sheet))
			{
				if (activeWorkbook.getActiveSheetIndex() != activeWorkbook.getSheetIndex(sheet))
				{
					activeWorkbook.setActiveSheet(activeWorkbook.getSheetIndex(sheet));
				}
			}
			else
			{
				result = addErrorMessage("sheet with name '" + sheetNode.asText() + "' does not exist");
			}
		}
		else if (sheetNode.isMissingNode())
		{
			JsonNode indexNode = getRequestNode().findPath(Key.INDEX.key());
			if (indexNode.isInt())
			{
				if (activeWorkbook.getNumberOfSheets() > indexNode.asInt())
				{
					if (activeWorkbook.getActiveSheetIndex() != indexNode.asInt())
					{
						activeWorkbook.setActiveSheet(indexNode.asInt());
					}
				}
				else
				{
					result = addErrorMessage(
							"sheet with " + Key.INDEX.key() + " " + indexNode.asInt() + " does not exist");
				}
			}
			else if (indexNode.isMissingNode())
			{
				result = addErrorMessage("missing argument '" + Key.SHEET.key() + "' or '" + Key.INDEX.key() + "'");
			}
			else
			{
				result = addErrorMessage("illegal argument '" + Key.INDEX.key() + "'");
			}
		}
		else
		{
			result = addErrorMessage("illegal argument '" + Key.SHEET.key() + "'");
		}
		if (result)
		{
			getResponseNode().put(Key.INDEX.key(), activeWorkbook.getActiveSheetIndex());
			getResponseNode().put(Key.SHEET.key(),
					activeWorkbook.getSheetAt(activeWorkbook.getActiveSheetIndex()).getSheetName());
		}
		return result;
	}

	private boolean doActiveSheetPresent()
	{
		boolean result = true;
		try
		{
			Sheet sheet = activeWorkbook.getSheetAt(activeWorkbook.getActiveSheetIndex());
			result = Objects.nonNull(sheet);
			getResponseNode().put(Key.INDEX.key(), result ? 1 : 0);
			getResponseNode().put(Key.SHEET.key(), result ? sheet.getSheetName() : "");
		}
		catch (IllegalArgumentException e)
		{
			result = false;
			getResponseNode().put(Key.INDEX.key(), 0);
			getResponseNode().put(Key.SHEET.key(), "");
		}
		return result;
	}

	private boolean doApplyCellStyles()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			CellRangeAddress cellRangeAddress = null;
			JsonNode cellNode = getRequestNode().findPath(Key.CELL.key());
			if (!cellNode.isMissingNode())
			{
				CellAddress cellAddress = getCellAddress(cellNode);
				cellRangeAddress = new CellRangeAddress(cellAddress.getRow(), cellAddress.getRow(),
						cellAddress.getColumn(), cellAddress.getColumn());
			}
			else
			{
				JsonNode rangeNode = getRequestNode().findPath(Key.RANGE.key());
				cellRangeAddress = getCellRangeAddress(rangeNode);
			}
			if (Objects.nonNull(cellRangeAddress))
			{
				Iterator<CellAddress> cellAddresses = cellRangeAddress.iterator();
				while (cellAddresses.hasNext())
				{
					CellAddress cellAddress = cellAddresses.next();
					Cell cell = getOrCreateCell(sheet, cellAddress);
					if (Objects.nonNull(cell))
					{
						MergedCellStyle m = new MergedCellStyle(cell.getCellStyle());
						result = m.applyRequestedStyles(getRequestNode(), getResponseNode());
						if (result)
						{
							CellStyle cellStyle = getCellStyle(sheet, m);
							cell.setCellStyle(cellStyle);
						}
						else
						{
							break;
						}
					}
					else
					{
						break;
					}
				}
			}
		}
		return result;
	}

	private boolean doApplyFontStyles()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			CellRangeAddress cellRangeAddress = null;
			JsonNode cellNode = getRequestNode().findPath(Key.CELL.key());
			if (!cellNode.isMissingNode())
			{
				CellAddress cellAddress = getCellAddress(cellNode);
				cellRangeAddress = new CellRangeAddress(cellAddress.getRow(), cellAddress.getRow(),
						cellAddress.getColumn(), cellAddress.getColumn());
			}
			else
			{
				JsonNode rangeNode = getRequestNode().findPath(Key.RANGE.key());
				cellRangeAddress = getCellRangeAddress(rangeNode);
			}
			if (Objects.nonNull(cellRangeAddress))
			{
				Iterator<CellAddress> cellAddresses = cellRangeAddress.iterator();
				while (cellAddresses.hasNext())
				{
					CellAddress cellAddress = cellAddresses.next();
					Row row = sheet.getRow(cellAddress.getRow());
					if (Objects.nonNull(row))
					{
						Cell cell = row.getCell(cellAddress.getColumn());
						CellStyle cellStyle = cell.getCellStyle();
						MergedCellStyle mcs = new MergedCellStyle(cellStyle);
						int fontIndex = cellStyle.getFontIndex();
						Font font = sheet.getWorkbook().getFontAt(fontIndex);
						MergedFont m = new MergedFont(font);
						if (m.applyRequestedFontStyles(getRequestNode(), getResponseNode()))
						{
							font = getFont(sheet, m);
							if (font.getIndex() != fontIndex)
							{
								mcs.setFontIndex(font.getIndex());
								cellStyle = getCellStyle(sheet, mcs);
								cell.setCellStyle(cellStyle);
							}
						}
						else
						{
							break;
						}
					}
					else
					{
						break;
					}
				}
			}
		}
		return result;
	}

	private boolean doAutoSizeColumns()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			CellRangeAddress cellRangeAddress = null;
			JsonNode cellNode = getRequestNode().findPath(Key.CELL.key());
			if (cellNode.isMissingNode())
			{
				JsonNode rangeNode = getRequestNode().findPath(Key.RANGE.key());
				if (!rangeNode.isMissingNode())
				{
					cellRangeAddress = getCellRangeAddress(rangeNode);
				}
			}
			else
			{
				CellAddress cellAddress = getCellAddress(cellNode);
				cellRangeAddress = new CellRangeAddress(cellAddress.getRow(), cellAddress.getColumn(),
						cellAddress.getRow(), cellAddress.getColumn());
			}
			if (Objects.nonNull(cellRangeAddress))
			{
				int leftCol = cellRangeAddress.getFirstColumn();
				int rightCol = cellRangeAddress.getLastColumn();
				for (int colIndex = leftCol; colIndex <= rightCol; colIndex++)
				{
					sheet.autoSizeColumn(colIndex);
				}
			}
		}
		return result;
	}

	private boolean doCopyCells()
	{
		boolean result = true;
		Sheet sourceSheet = getSheet(getRequestNode());
		Sheet targetSheet = sourceSheet;
		CellRangeAddress sourceCellRangeAddress = null;
		CellRangeAddress targetCellRangeAddress = null;
		JsonNode sourceNode = getRequestNode().findPath(Key.SOURCE.key());
		if (sourceNode.isTextual())
		{
			sourceCellRangeAddress = getCellRangeAddress(sourceNode);
			if (Objects.nonNull(sourceCellRangeAddress))
			{
				JsonNode targetNode = getRequestNode().findPath(Key.TARGET.key());
				if (targetNode.isTextual())
				{
					targetCellRangeAddress = getCellRangeAddress(targetNode);
					if (Objects.isNull(targetCellRangeAddress))
					{
						result = addErrorMessage("illegal argument '" + targetNode.asText() + "'");
					}
				}
				else if (targetNode.isObject())
				{
					targetSheet = getSheet(ObjectNode.class.cast(targetNode));
					targetCellRangeAddress = getCellRangeAddress(targetNode);
				}
				else
				{
					result = addErrorMessage("illegal argument '" + Key.TARGET.key() + "'");
				}
			}
			else
			{
				result = addErrorMessage("illegal argument '" + sourceNode.asText() + "'");
			}
		}
		else if (sourceNode.isObject())
		{
			sourceSheet = getSheet(ObjectNode.class.cast(sourceNode));
			sourceCellRangeAddress = getCellRangeAddress(sourceNode);
			JsonNode targetNode = getRequestNode().findPath(Key.TARGET.key());
			if (targetNode.isTextual())
			{
				targetCellRangeAddress = getCellRangeAddress(targetNode);
				if (Objects.isNull(targetCellRangeAddress))
				{
					result = addErrorMessage("illegal argument '" + targetNode.asText() + "'");
				}
			}
			else if (targetNode.isObject())
			{
				targetSheet = getSheet(ObjectNode.class.cast(targetNode));
				targetCellRangeAddress = getCellRangeAddress(targetNode);
			}
			else
			{
				result = addErrorMessage("illegal argument '" + Key.TARGET.key() + "'");
			}
		}
		else
		{
			result = addErrorMessage("illegal argument '" + Key.SOURCE.key() + "'");
		}
		if (Objects.nonNull(sourceCellRangeAddress) && Objects.nonNull(targetCellRangeAddress))
		{
			if (sourceSheet == targetSheet)
			{
				if (sourceCellRangeAddress.intersects(targetCellRangeAddress))
				{
					result = addErrorMessage("source range and target range must not intersect");
				}
			}
			if (result)
			{
				if (sourceCellRangeAddress.getNumberOfCells() == 1)
				{
					Row sourceRow = sourceSheet.getRow(sourceCellRangeAddress.getFirstRow());
					if (Objects.nonNull(sourceRow))
					{
						Cell sourceCell = sourceRow.getCell(sourceCellRangeAddress.getFirstColumn());
						if (Objects.nonNull(sourceCell))
						{
							Iterator<CellAddress> targetAddresses = targetCellRangeAddress.iterator();
							while (targetAddresses.hasNext())
							{
								CellAddress sourceAddress = new CellAddress(sourceCell);
								CellAddress targetAddress = targetAddresses.next();
								int rowDiff = targetAddress.getRow() - sourceAddress.getRow();
								int colDiff = targetAddress.getColumn() - sourceAddress.getColumn();
								if (sourceCell.getCellType().equals(CellType.FORMULA))
								{
									String copiedFormula = copyFormula(sourceSheet, sourceCell.getCellFormula(),
											rowDiff, colDiff);
									Cell targetCell = getOrCreateCell(targetSheet, targetAddress);
									targetCell.setCellFormula(copiedFormula);
								}
								else
								{
									int targetTop = targetCellRangeAddress.getFirstRow();
									int targetBottom = targetCellRangeAddress.getLastRow();
									int targetLeft = targetCellRangeAddress.getFirstColumn();
									int targetRight = targetCellRangeAddress.getLastColumn();
									for (int r = targetTop; r <= targetBottom; r++)
									{
										Row targetRow = getOrCreateRow(targetSheet, r);
										for (int cell = targetLeft; cell <= targetRight; cell++)
										{
											Cell targetCell = getOrCreateCell(targetRow, cell);
											if (sourceCell.getCellType().equals(CellType.STRING))
											{
												targetCell.setCellValue(sourceCell.getRichStringCellValue());

											}
											else if (sourceCell.getCellType().equals(CellType.NUMERIC))
											{
												targetCell.setCellValue(sourceCell.getNumericCellValue());
											}
										}
									}
								}
							}
						}
					}
				}
				else if (sourceCellRangeAddress.getNumberOfCells() == targetCellRangeAddress.getNumberOfCells()
						&& sourceCellRangeAddress.getLastRow()
								- sourceCellRangeAddress.getFirstRow() == targetCellRangeAddress.getLastRow()
										- targetCellRangeAddress.getFirstRow())
				{
					Iterator<CellAddress> sourceAddresses = sourceCellRangeAddress.iterator();
					Iterator<CellAddress> targetAddresses = targetCellRangeAddress.iterator();
					while (sourceAddresses.hasNext())
					{
						CellAddress sourceAddress = sourceAddresses.next();
						Row sourceRow = sourceSheet.getRow(sourceAddress.getRow());
						if (Objects.nonNull(sourceRow))
						{
							Cell sourceCell = sourceRow.getCell(sourceAddress.getColumn());
							if (Objects.nonNull(sourceCell))
							{
								CellAddress targetAddress = targetAddresses.next();
								Row targetRow = getOrCreateRow(targetSheet, targetAddress.getRow());
								Cell targetCell = getOrCreateCell(targetRow, targetAddress.getColumn());
								copyCell(sourceCell, targetCell);
							}
						}
					}
				}
				else
				{
					result = addErrorMessage("source and target range dimensions must not differ");
				}
			}
			else
			{
				result = addErrorMessage("missing argument 'sheet' for source");
			}
		}
		else
		{
			result = addErrorMessage("missing argument 'target'");
		}
		return result;
	}

	private boolean doCreateSheet()
	{
		boolean result = true;
		Sheet sheet = null;
		JsonNode sheetNode = getRequestNode().findPath(Key.SHEET.key());
		if (sheetNode.isMissingNode())
		{
			sheet = activeWorkbook.createSheet();
		}
		else if (sheetNode.isTextual())
		{
			try
			{
				sheet = activeWorkbook.createSheet(sheetNode.asText());
			}
			catch (IllegalArgumentException e)
			{
				result = addErrorMessage("illegal argument 'sheet' ('" + sheetNode.asText() + "' already exists)");
			}
		}
		if (result)
		{
			getResponseNode().put(Key.SHEET.key(), sheet.getSheetName());
			getResponseNode().put(Key.INDEX.key(), activeWorkbook.getSheetIndex(sheet));
		}
		return result;
	}

	private boolean doCreateWorkbook()
	{
		boolean result = true;
		Type type = null;
		JsonNode typeNode = getRequestNode().findPath(Key.TYPE.key());
		if (typeNode.isTextual())
		{
			type = Type.findByExtension(typeNode.asText());
		}
		else if (typeNode.isMissingNode())
		{
			type = Type.XLSX;
		}
		if (Objects.nonNull(type))
		{
			switch (type)
			{
				case XLSX:
				{
					activeWorkbook = new XSSFWorkbook();
				}
				case XLS:
				{
					activeWorkbook = new HSSFWorkbook();
				}
			}
		}
		else
		{
			result = addErrorMessage("illegal extension '" + typeNode.asText() + "'");
		}
		return result;
	}

	private boolean doDropSheet()
	{
		boolean result = true;
		JsonNode sheetNode = getRequestNode().findPath(Key.SHEET.key());
		if (sheetNode.isTextual())
		{
			Sheet sheet = activeWorkbook.getSheet(sheetNode.asText());
			if (Objects.nonNull(sheet))
			{
				activeWorkbook.removeSheetAt(activeWorkbook.getSheetIndex(sheet));
			}
			else
			{
				result = addErrorMessage("sheet with name '" + sheetNode.asText() + "' does not exist");
			}
		}
		else if (sheetNode.isMissingNode())
		{
			JsonNode indexNode = getRequestNode().findPath(Key.INDEX.key());
			if (indexNode.isInt())
			{
				if (activeWorkbook.getActiveSheetIndex() > -1)
				{
					activeWorkbook.removeSheetAt(indexNode.asInt());
				}
				else
				{
					result = addErrorMessage(
							"sheet with " + Key.INDEX.key() + " " + indexNode.asInt() + " does not exist");
				}
			}
			else if (indexNode.isMissingNode())
			{
				if (activeWorkbook.getNumberOfSheets() > 0)
				{
					if (activeWorkbook.getActiveSheetIndex() > -1)
					{
						activeWorkbook.removeSheetAt(activeWorkbook.getActiveSheetIndex());
					}
				}
				else
				{
					result = addErrorMessage("there is no active sheet present");
				}
			}
		}
		return result;
	}

	private boolean doGetActiveSheet()
	{
		boolean result = true;
		Sheet sheet = activeWorkbook.getSheetAt(activeWorkbook.getActiveSheetIndex());
		if (Objects.nonNull(sheet))
		{
			getResponseNode().put(Key.INDEX.key(), activeWorkbook.getSheetIndex(sheet));
			getResponseNode().put(Key.SHEET.key(), sheet.getSheetName());
		}
		else
		{
			result = addErrorMessage("there is no active sheet present");
		}
		return result;
	}

	private boolean doGetSheetNames()
	{
		boolean result = true;
		ArrayNode sheetsNode = getResponseNode().arrayNode();
		ArrayNode indexNode = getResponseNode().arrayNode();
		int numberOfSheets = activeWorkbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++)
		{
			sheetsNode.add(activeWorkbook.getSheetAt(i).getSheetName());
			indexNode.add(i);
		}
		getResponseNode().set(Key.SHEET.key(), sheetsNode);
		getResponseNode().set(Key.INDEX.key(), indexNode);
		return result;
	}

	private boolean doMoveSheet()
	{
		boolean result = true;
		Sheet sheet = activeWorkbook.getSheetAt(activeWorkbook.getActiveSheetIndex());
		JsonNode sourceNode = getRequestNode().findPath(Key.SOURCE.key());
		if (sourceNode.isTextual())
		{
			sheet = activeWorkbook.getSheet(sourceNode.asText());
		}
		else if (sourceNode.isInt())
		{
			sheet = activeWorkbook.getSheetAt(sourceNode.asInt());
		}
		else if (!sourceNode.isMissingNode())
		{
			result = addErrorMessage("illegal argument '" + Key.SOURCE.key() + "'");
		}
		if (result)
		{
			JsonNode targetNode = getRequestNode().findPath(Key.TARGET.key());
			if (targetNode.isInt())
			{
				if (activeWorkbook.getNumberOfSheets() >= targetNode.asInt())
				{
					if (targetNode.asInt() < 0)
					{
						result = addErrorMessage("illegal argument '" + Key.TARGET.key()
								+ "' (sheet index is out of range: " + targetNode.asInt() + ")");
					}

					if (activeWorkbook.getActiveSheetIndex() != targetNode.asInt())
					{
						activeWorkbook.setSheetOrder(sheet.getSheetName(), targetNode.asInt());
					}
				}
				else
				{
					result = addErrorMessage(
							"illegal argument '" + Key.TARGET.key() + "' (sheet index is out of range: "
									+ targetNode.asInt() + " > " + activeWorkbook.getNumberOfSheets() + ")");
				}
			}
			else if (targetNode.isMissingNode())
			{
				result = addErrorMessage("missing argument '" + Key.TARGET.key() + "'");
			}
		}
		return result;
	}

	private boolean doReleaseWorkbook()
	{
		activeWorkbook = null;
		return true;
	}

	private boolean doRenameSheet()
	{
		boolean result = true;
		JsonNode sheetNode = getRequestNode().findPath(Key.SHEET.key());
		if (sheetNode.isTextual())
		{
			int index = activeWorkbook.getActiveSheetIndex();
			JsonNode indexNode = getRequestNode().findPath(Key.INDEX.key());
			if (indexNode.isInt())
			{
				index = indexNode.asInt();
			}
			else if (!indexNode.isMissingNode())
			{
				result = addErrorMessage("illegal argument '" + Key.INDEX.key() + "'");
			}
			if (result)
			{
				activeWorkbook.setSheetName(index, sheetNode.asText());
			}
		}
		else if (sheetNode.isMissingNode())
		{
			result = addErrorMessage("missing argument '" + Key.SHEET.key() + "'");
		}
		else
		{
			result = addErrorMessage("illegal argument '" + Key.SHEET.key() + "'");
		}
		return result;
	}

	private boolean doRotateCells()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			CellRangeAddress cellRangeAddress = null;
			JsonNode cellNode = getRequestNode().findPath(Key.CELL.key());
			if (!cellNode.isMissingNode())
			{
				CellAddress cellAddress = getCellAddress(cellNode);
				cellRangeAddress = new CellRangeAddress(cellAddress.getRow(), cellAddress.getRow(),
						cellAddress.getColumn(), cellAddress.getColumn());
			}
			else
			{
				JsonNode rangeNode = getRequestNode().findPath(Key.RANGE.key());
				cellRangeAddress = getCellRangeAddress(rangeNode);
			}
			if (Objects.nonNull(cellRangeAddress))
			{
				int rotation = Integer.MIN_VALUE;
				JsonNode rotationNode = getRequestNode().findPath(Key.ROTATION.key());
				if (rotationNode.isInt())
				{
					rotation = IntNode.class.cast(getRequestNode().get(Key.ROTATION.key())).asInt();
				}
				else if (rotationNode.isMissingNode())
				{
					result = addErrorMessage("missing argument '" + Key.ROTATION.key() + "'");
				}
				else
				{
					result = addErrorMessage("illegal argument '" + Key.ROTATION.key() + "'");
				}
				if (rotation != Integer.MIN_VALUE)
				{
					Iterator<CellAddress> cellAddresses = cellRangeAddress.iterator();
					while (cellAddresses.hasNext())
					{
						CellAddress cellAddress = cellAddresses.next();
						Row row = sheet.getRow(cellAddress.getRow());
						if (Objects.nonNull(row))
						{
							Cell cell = row.getCell(cellAddress.getColumn());
							if (Objects.nonNull(cell))
							{
								CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
								cellStyle.setRotation((short) rotation);
								cell.setCellStyle(cellStyle);
							}
						}
					}
				}
			}
		}
		return result;
	}

	private boolean doSaveWorkbook()
	{
		boolean result = true;
		JsonNode pathNode = getRequestNode().findPath(Key.PATH.key());
		if (pathNode.isTextual())
		{
			File file = new File(pathNode.asText());
			OutputStream os = null;
			try
			{
				file.getCanonicalPath();
				if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls"))
				{
					if (XSSFWorkbook.class.isInstance(activeWorkbook))
					{
						file = new File(file.getAbsolutePath() + ".xlsx");
					}
					else if (HSSFWorkbook.class.isInstance(activeWorkbook))
					{
						file = new File(file.getAbsolutePath() + ".xls");
					}
				}
				os = new FileOutputStream(file);
				activeWorkbook.write(os);
			}
			catch (Exception e)
			{
				result = addErrorMessage("saving workbook failed (" + e.getLocalizedMessage() + ")");
			}
			finally
			{
				if (Objects.nonNull(os))
				{
					try
					{
						os.flush();
						os.close();
					}
					catch (Exception e)
					{
					}
				}
			}
		}
		else if (pathNode.isMissingNode())
		{
			result = addErrorMessage("missing argument '" + Key.PATH.key() + "'");
		}
		else
		{
			result = addErrorMessage("illegal argument '" + Key.PATH.key() + "'");
		}
		return result;
	}

	private boolean doSetCell(Sheet sheet, CellAddress cellAddress, JsonNode valueNode)
	{
		boolean result = true;
		Cell cell = getOrCreateCell(sheet, cellAddress);
		if (valueNode.isNumber())
		{
			cell.setCellValue(valueNode.asDouble());
		}
		else if (valueNode.isTextual())
		{
			if (!valueNode.asText().trim().isEmpty())
			{
				try
				{
					cell.setCellValue(DateUtil.parseDateTime(valueNode.asText()));
					MergedCellStyle mcs = new MergedCellStyle(cell.getCellStyle());
					int formatIndex = BuiltinFormats.getBuiltinFormat("m/d/yy");
					mcs.setDataFormat((short) formatIndex);
					CellStyle cellStyle = getCellStyle(sheet, mcs);
					cell.setCellStyle(cellStyle);
				}
				catch (Exception tpe)
				{
					try
					{
						Date date = DateFormat.getDateInstance().parse(valueNode.asText());
						cell.setCellValue(DateUtil.getExcelDate(date));
						MergedCellStyle mcs = new MergedCellStyle(cell.getCellStyle());
						int formatIndex = BuiltinFormats.getBuiltinFormat("m/d/yy");
						mcs.setDataFormat((short) formatIndex);
						CellStyle cellStyle = getCellStyle(sheet, mcs);
						cell.setCellStyle(cellStyle);
					}
					catch (ParseException dpe)
					{
						try
						{
							double time = DateUtil.convertTime(valueNode.asText());
							cell.setCellValue(time);
							MergedCellStyle mcs = new MergedCellStyle(cell.getCellStyle());
							int formatIndex = BuiltinFormats.getBuiltinFormat("h:mm");
							mcs.setDataFormat((short) formatIndex);
							CellStyle cellStyle = getCellStyle(sheet, mcs);
							cell.setCellStyle(cellStyle);
						}
						catch (Exception e)
						{
							try
							{
								cell.setCellFormula(valueNode.asText());
								FormulaEvaluator evaluator = activeWorkbook.getCreationHelper()
										.createFormulaEvaluator();
								CellType cellType = evaluator.evaluateFormulaCell(cell);
								System.out.println(cellType);
							}
							catch (FormulaParseException fpe)
							{
								setRichTextString(cell, valueNode.asText());
							}
						}
					}
				}
			}
		}
		else if (valueNode.isBoolean())
		{
			cell.setCellValue(valueNode.asBoolean());
		}
		else
		{
			// TODO Other types?
			System.out.println();
		}
		return result;
	}

	private boolean doSetCells()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		result = Objects.nonNull(sheet);
		if (result)
		{
			JsonNode cellNode = getRequestNode().findPath(Key.CELL.key());
			JsonNode valuesNode = getRequestNode().findPath(Key.VALUES.key());
			if (cellNode.isMissingNode())
			{
				result = addErrorMessage("missing argument '" + Key.CELL.key() + "'");
			}
			else if (valuesNode.isMissingNode())
			{
				result = addErrorMessage("missing argument '" + Key.VALUES.key() + "'");
			}
			else if (valuesNode.isArray())
			{
				ArrayNode valuesArrayNode = ArrayNode.class.cast(valuesNode);
				if (cellNode.isArray())
				{
					if (cellNode.size() == valuesArrayNode.size())
					{
						result = setCells(sheet, ArrayNode.class.cast(cellNode), valuesArrayNode);
					}
					else
					{
						result = addErrorMessage("size of 'cell' array does not equal to size of 'values' array");
					}
				}
				else
				{
					Direction direction = Direction.DEFAULT;
					JsonNode directionNode = getRequestNode().findPath(Key.DIRECTION.key());
					if (directionNode.isTextual())
					{
						try
						{
							direction = Direction.valueOf(directionNode.asText().toUpperCase());
						}
						catch (Exception e)
						{
							result = addErrorMessage("invalid argument 'direction'");
						}
					}
					else if (directionNode.isMissingNode())
					{
					}
					else
					{
						result = addErrorMessage("invalid argument 'direction'");
					}
					if (result)
					{
						if (cellNode.isTextual())
						{
							doSetCells(sheet, TextNode.class.cast(cellNode), valuesArrayNode, direction);
						}
						else if (cellNode.isObject())
						{
							doSetCells(sheet, ObjectNode.class.cast(cellNode), valuesArrayNode, direction);
						}
						else
						{
							result = addErrorMessage("invalid argument '" + Key.CELL.key() + "'");
						}
					}
				}
			}
			else
			{
				result = addErrorMessage("invalid argument '" + Key.VALUES.key() + "'");
			}
		}
		else
		{
			result = addErrorMessage("missing_argument '" + Key.SHEET.key() + "'");
		}
		return result;
	}

	private boolean doSetCells(Sheet sheet, CellAddress cellAddress, ArrayNode valuesNode, Direction direction)
	{
		boolean result = Objects.nonNull(cellAddress);
		if (result)
		{
			if (valuesNode.size() > 0)
			{
				if (direction.validRange(getResponseNode(), sheet.getWorkbook(), cellAddress, valuesNode.size()))
				{
					for (int i = 0; i < valuesNode.size(); i++)
					{
						JsonNode valueNode = valuesNode.get(i);
						result = doSetCell(sheet, cellAddress, valueNode);
						if (result)
						{
							cellAddress = direction.nextIndex(cellAddress);
						}
						else
						{
							break;
						}
					}
				}
			}
			else
			{
				result = addErrorMessage("invalid_argument '" + Key.VALUES.key() + "'");
			}
		}
		return result;
	}

	private boolean doSetCells(Sheet sheet, ObjectNode cellNode, ArrayNode valuesNode, Direction direction)
	{
		boolean result = true;
		try
		{
			CellAddress cellAddress = getCellAddress(cellNode);
			result = doSetCells(sheet, cellAddress, valuesNode, direction);
		}
		catch (Exception e)
		{
			result = addErrorMessage("invalid argument '" + cellNode.asText() + "'");
		}
		return result;
	}

	private boolean doSetCells(Sheet sheet, TextNode cellNode, ArrayNode valuesNode, Direction direction)
	{
		boolean result = true;
		try
		{
			CellAddress cellAddress = getCellAddress(cellNode);
			result = doSetCells(sheet, cellAddress, valuesNode, direction);
		}
		catch (Exception e)
		{
			result = addErrorMessage("invalid argument '" + cellNode.asText() + "'");
		}
		return result;
	}

	private boolean doSetFooters()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			Footer footer = sheet.getFooter();
			JsonNode leftNode = getRequestNode().findPath(Key.LEFT.key());
			if (leftNode.isTextual())
			{
				footer.setLeft(leftNode.asText());

			}
			JsonNode centerNode = getRequestNode().findPath(Key.CENTER.key());
			if (centerNode.isTextual())
			{
				footer.setCenter(centerNode.asText());

			}
			JsonNode rightNode = getRequestNode().findPath(Key.RIGHT.key());
			if (rightNode.isTextual())
			{
				footer.setRight(rightNode.asText());
			}
		}
		return result;
	}

	private boolean doSetHeaders()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			Header header = sheet.getHeader();
			JsonNode leftNode = getRequestNode().findPath(Key.LEFT.key());
			if (leftNode.isTextual())
			{
				header.setLeft(leftNode.asText());

			}
			JsonNode centerNode = getRequestNode().findPath(Key.CENTER.key());
			if (centerNode.isTextual())
			{
				header.setCenter(centerNode.asText());

			}
			JsonNode rightNode = getRequestNode().findPath(Key.RIGHT.key());
			if (rightNode.isTextual())
			{
				header.setRight(rightNode.asText());
			}
		}
		return result;
	}

	private boolean doSetPrintSetup()
	{
		boolean result = true;
		Sheet sheet = getSheet(getRequestNode());
		if (Objects.nonNull(sheet))
		{
			JsonNode orientationNode = getRequestNode().findPath(Key.ORIENTATION.key());
			if (!orientationNode.isMissingNode())
			{
				PrintOrientation orientation = PrintOrientation.DEFAULT;
				if (orientationNode.isTextual())
				{
					try
					{
						orientation = PrintOrientation.valueOf(orientationNode.asText().toUpperCase());
					}
					catch (Exception e)
					{
						result = addErrorMessage("invalid argument 'orientation'");
					}
				}
				switch (orientation)
				{
					case LANDSCAPE:
					{
						sheet.getPrintSetup().setLandscape(true);
					}
					case PORTRAIT:
					{
						sheet.getPrintSetup().setNoOrientation(false);
						sheet.getPrintSetup().setLandscape(false);
					}
					default:
					{
						sheet.getPrintSetup().setNoOrientation(true);
					}
				}
			}
			int copies = 1;
			JsonNode copiesNode = getRequestNode().findPath(Key.COPIES.key());
			if (!copiesNode.isMissingNode())
			{
				if (copiesNode.isInt())
				{
					copies = copiesNode.asInt();
				}
				if (copies > 0 && copies <= Short.MAX_VALUE)
				{
					sheet.getPrintSetup().setCopies((short) copies);
				}
			}
		}
		return result;
	}

	private CellAddress getCellAddress(JsonNode cellNode)
	{
		CellAddress cellAddress = null;
		if (!cellNode.isMissingNode())
		{
			if (cellNode.isTextual())
			{
				cellAddress = new CellAddress(cellNode.asText());
				if (Objects.isNull(cellAddress))
				{
					addErrorMessage("illegal argument '" + cellNode.asText() + "'");
				}
			}
			else if (cellNode.isObject())
			{
				JsonNode rowNode = cellNode.findPath(Key.ROW.key());
				if (rowNode.isInt())
				{
					JsonNode columnNode = cellNode.findPath(Key.COL.key());
					if (columnNode.isInt())
					{
						cellAddress = new CellAddress(rowNode.asInt(), columnNode.asInt());
					}
					else if (columnNode.isMissingNode())
					{
						addErrorMessage("missing argument '" + Key.COL.key() + "'");
					}
					else
					{
						addErrorMessage("illegal argument '" + Key.COL.key() + "'");
					}
				}
				else if (rowNode.isMissingNode())
				{
					addErrorMessage("missing argument '" + Key.ROW.key() + "'");
				}
				else
				{
					addErrorMessage("illegal argument '" + Key.ROW.key() + "'");
				}
			}
			else
			{
				addErrorMessage("illegal argument '" + Key.CELL.key() + "'");
			}
		}
		return cellAddress;
	}

	private CellAddress getCellAddress(JsonNode cellNode, String key)
	{
		CellAddress cellAddress = null;
		if (cellNode.isTextual())
		{
			cellAddress = new CellAddress(cellNode.asText());
			if (Objects.isNull(cellAddress))
			{
				addErrorMessage("illegal argument '" + cellNode.asText() + "'");
			}
		}
		else if (cellNode.isObject())
		{
			cellAddress = getCellAddress(ObjectNode.class.cast(cellNode));
		}
		else if (Objects.isNull(cellNode) || cellNode.isMissingNode())
		{
			addErrorMessage("missing argument '" + key + "'");
		}
		else
		{
			addErrorMessage("invalid argument '" + key + "'");
		}
		return cellAddress;
	}

	private CellRangeAddress getCellRangeAddress(JsonNode rangeNode)
	{
		CellRangeAddress cellRangeAddress = null;
		if (!rangeNode.isMissingNode())
		{
			CellAddress topLeftAddress = null;
			CellAddress bottomRightAddress = null;
			if (rangeNode.isTextual())
			{
				String range = rangeNode.asText();
				String[] rangeParts = range.split(":");
				String topLeftCell = null;
				String bottomRightCell = null;
				if (rangeParts.length > 0)
				{
					topLeftCell = rangeParts[0];
					topLeftAddress = new CellAddress(topLeftCell);
					bottomRightAddress = new CellAddress(topLeftCell);
				}
				if (rangeParts.length == 2)
				{
					bottomRightCell = rangeParts[1];
					bottomRightAddress = new CellAddress(bottomRightCell);
				}
				if (Objects.nonNull(topLeftAddress) && Objects.nonNull(bottomRightAddress))
				{
					cellRangeAddress = new CellRangeAddress(topLeftAddress.getRow(), bottomRightAddress.getRow(),
							topLeftAddress.getColumn(), bottomRightAddress.getColumn());
				}
			}
			else if (rangeNode.isObject())
			{
				topLeftAddress = getCellAddress(rangeNode.findPath(Key.TOP_LEFT.key()));
				if (Objects.nonNull(topLeftAddress))
				{
					bottomRightAddress = getCellAddress(rangeNode.findPath(Key.BOTTOM_RIGHT.key()));
					if (Objects.nonNull(bottomRightAddress))
					{
						cellRangeAddress = new CellRangeAddress(topLeftAddress.getRow(), bottomRightAddress.getRow(),
								topLeftAddress.getColumn(), bottomRightAddress.getColumn());
					}
					else
					{
						JsonNode topNode = rangeNode.findPath(Key.TOP.key());
						if (topNode.isInt())
						{
							JsonNode leftNode = rangeNode.get(Key.LEFT.key());
							if (leftNode.isInt())
							{
								topLeftAddress = new CellAddress(topNode.asInt(), leftNode.asInt());
							}
							else
							{
								addErrorMessage("illegal argument '" + Key.LEFT.key() + "'");
							}
						}
						else
						{
							addErrorMessage("illegal argument '" + Key.TOP.key() + "'");
						}
					}
				}
				else
				{
					JsonNode topNode = rangeNode.findPath(Key.TOP.key());
					if (topNode.isInt())
					{
						JsonNode leftNode = rangeNode.get(Key.LEFT.key());
						if (leftNode.isInt())
						{
							topLeftAddress = new CellAddress(topNode.asInt(), leftNode.asInt());
							if (Objects.nonNull(topLeftAddress))
							{
								JsonNode bottomNode = rangeNode.findPath(Key.TOP.key());
								if (bottomNode.isInt())
								{
									JsonNode rightNode = rangeNode.get(Key.LEFT.key());
									if (rightNode.isInt())
									{
										bottomRightAddress = new CellAddress(bottomNode.asInt(), rightNode.asInt());
										if (Objects.isNull(bottomRightAddress))
										{
											bottomRightAddress = getCellAddress(
													rangeNode.findPath(Key.BOTTOM_RIGHT.key()));
											if (Objects.nonNull(bottomRightAddress))
											{
												cellRangeAddress = new CellRangeAddress(topLeftAddress.getRow(),
														bottomRightAddress.getRow(), topLeftAddress.getColumn(),
														bottomRightAddress.getColumn());
											}
											else
											{
												addErrorMessage("illegal argument '" + Key.BOTTOM_RIGHT.key() + "'");
											}
										}
										else
										{
											cellRangeAddress = new CellRangeAddress(topLeftAddress.getRow(),
													bottomRightAddress.getRow(), topLeftAddress.getColumn(),
													bottomRightAddress.getColumn());
										}
									}
									else
									{
										addErrorMessage("illegal argument '" + Key.LEFT.key() + "'");
									}
								}
								else
								{
									addErrorMessage("illegal argument '" + Key.TOP.key() + "'");
								}
							}
						}
						else
						{
							addErrorMessage("illegal argument '" + Key.LEFT.key() + "'");
						}
					}
					else
					{
						addErrorMessage("illegal argument '" + Key.TOP.key() + "'");
					}
				}
			}
			else
			{
				addErrorMessage("illegal argument '" + Key.RANGE.key() + "'");
			}
		}
		return cellRangeAddress;
	}

	private CellStyle getCellStyle(Sheet sheet, MergedCellStyle m)
	{
		CellStyle cellStyle = null;
		for (int i = 0; i < sheet.getWorkbook().getNumCellStyles(); i++)
		{
			CellStyle cs = sheet.getWorkbook().getCellStyleAt(i);
			if (cs.getAlignment().equals(m.getHalign()) && cs.getVerticalAlignment().equals(m.getValign())
					&& cs.getBorderBottom().equals(m.getBottom()) && cs.getBorderLeft().equals(m.getLeft())
					&& cs.getBorderRight().equals(m.getRight()) && cs.getBorderTop().equals(m.getTop())
					&& cs.getBottomBorderColor() == m.getbColor() && cs.getLeftBorderColor() == m.getlColor()
					&& cs.getRightBorderColor() == m.getrColor() && cs.getTopBorderColor() == m.gettColor()
					&& cs.getDataFormat() == m.getDataFormat() && cs.getFillBackgroundColor() == m.getBgColor()
					&& cs.getFillForegroundColor() == m.getFgColor() && cs.getFontIndex() == m.getFontIndex()
					&& cs.getFillPattern().equals(m.getFillPattern()) && cs.getShrinkToFit() == m.getShrinkToFit()
					&& cs.getWrapText() == m.getWrapText())
			{
				cellStyle = cs;
				break;
			}
		}
		if (Objects.isNull(cellStyle))
		{
			cellStyle = sheet.getWorkbook().createCellStyle();
			m.applyToCellStyle(sheet, cellStyle);
		}
		return cellStyle;
	}

	private Font getFont(Sheet sheet, MergedFont m)
	{
		Font font = null;
		for (int i = 0; i < sheet.getWorkbook().getNumberOfFonts(); i++)
		{
			Font f = sheet.getWorkbook().getFontAt(i);
			if (f.getFontName().equals(m.getName()) && f.getFontHeightInPoints() == m.getSize()
					&& f.getBold() == m.getBold() && f.getItalic() == m.getItalic()
					&& f.getUnderline() == m.getUnderline() && f.getStrikeout() == m.getStrikeOut()
					&& f.getTypeOffset() == m.getTypeOffset() && f.getColor() == m.getColor())
			{
				font = f;
				break;
			}
		}
		if (Objects.isNull(font))
		{
			font = sheet.getWorkbook().createFont();
			font.setFontName(m.getName());
			font.setFontHeightInPoints(m.getSize().shortValue());
			font.setBold(m.getBold().booleanValue());
			font.setItalic(m.getItalic().booleanValue());
			font.setUnderline(m.getUnderline().byteValue());
			font.setStrikeout(m.getStrikeOut().booleanValue());
			font.setTypeOffset(m.getTypeOffset().shortValue());
			font.setColor(m.getColor().shortValue());
		}
		return font;
	}

	private FormulaParsingWorkbook getFormulaParsingWorkbook(Sheet sheet)
	{
		FormulaParsingWorkbook workbookWrapper = null;
		if (XSSFSheet.class.isInstance(sheet))
		{
			workbookWrapper = XSSFEvaluationWorkbook.create(XSSFSheet.class.cast(sheet).getWorkbook());
		}
		else
		{
			workbookWrapper = HSSFEvaluationWorkbook.create(HSSFSheet.class.cast(sheet).getWorkbook());
		}
		return workbookWrapper;
	}

	private FormulaRenderingWorkbook getFormulaRenderingWorkbook(Sheet sheet)
	{
		FormulaRenderingWorkbook workbookWrapper = null;
		if (XSSFSheet.class.isInstance(sheet))
		{
			workbookWrapper = XSSFEvaluationWorkbook.create(XSSFSheet.class.cast(sheet).getWorkbook());
		}
		else
		{
			workbookWrapper = HSSFEvaluationWorkbook.create(HSSFSheet.class.cast(sheet).getWorkbook());
		}
		return workbookWrapper;
	}

	private Cell getOrCreateCell(Row row, int colIndex)
	{
		Cell cell = null;
		if (Objects.nonNull(row))
		{
			if (validateColIndex(colIndex))
			{
				cell = row.getCell(colIndex);
				if (Objects.isNull(cell))
				{
					cell = row.createCell(colIndex);
				}
			}
			else
			{
				addErrorMessage("illegal cell index (" + colIndex + " > "
						+ activeWorkbook.getSpreadsheetVersion().getLastColumnIndex() + ")");
			}
		}
		return cell;
	}

	private Cell getOrCreateCell(Sheet sheet, CellAddress cellAddress)
	{
		Cell cell = null;
		if (Objects.nonNull(cellAddress))
		{
			Row row = getOrCreateRow(sheet, cellAddress.getRow());
			if (Objects.nonNull(row))
			{
				cell = getOrCreateCell(row, cellAddress.getColumn());
			}
		}
		return cell;
	}

	private Row getOrCreateRow(Sheet sheet, int rowIndex)
	{
		Row row = null;
		if (validateRowIndex(rowIndex))
		{
			row = sheet.getRow(rowIndex);
			if (Objects.isNull(row))
			{
				row = sheet.createRow(rowIndex);
			}
		}
		else
		{
			addErrorMessage("illegal row index (" + rowIndex + " > "
					+ sheet.getWorkbook().getSpreadsheetVersion().getLastRowIndex() + ")");
		}
		return row;
	}

	private Sheet getSheet(ObjectNode parentNode)
	{
		Sheet sheet = null;
		try
		{
			sheet = activeWorkbook.getSheetAt(activeWorkbook.getActiveSheetIndex());
			JsonNode sheetNode = getRequestNode().findPath(Key.SHEET.key());
			if (sheetNode.isMissingNode())
			{
				JsonNode indexNode = getRequestNode().findPath(Key.INDEX.key());
				if (indexNode.isInt())
				{
					sheet = activeWorkbook.getSheetAt(indexNode.asInt());
					if (Objects.isNull(sheet))
					{
						addErrorMessage("sheet index with (" + indexNode.asInt() + ") does not exist");
					}
				}
				else if (!indexNode.isMissingNode())
				{
					addErrorMessage("illegal argument '" + Key.INDEX.key() + "'");
				}
			}
			else if (sheetNode.isTextual())
			{
				sheet = activeWorkbook.getSheet(sheetNode.asText());
				if (Objects.isNull(sheet))
				{
					addErrorMessage("sheet with name '" + sheetNode.asText() + "' does not exist");
				}
			}
			else
			{
				addErrorMessage("illegal argument '" + Key.SHEET.key() + "'");
			}
		}
		catch (IllegalArgumentException e)
		{
			sheet = null;
			addErrorMessage(e.getLocalizedMessage().toLowerCase());
		}
		return sheet;
	}

	private boolean setCells(Sheet sheet, ArrayNode cellNode, ArrayNode valuesNode)
	{
		boolean result = true;
		for (int i = 0; i < cellNode.size(); i++)
		{
			CellAddress cellAddress = getCellAddress(cellNode.get(i), Key.CELL.key());
			JsonNode valueNode = valuesNode.get(i);
			result = doSetCell(sheet, cellAddress, valueNode);

		}
		return result;
	}

	private void setRichTextString(Cell cell, String value)
	{
		if (XSSFCell.class.isInstance(cell))
		{
			cell.setCellValue(new XSSFRichTextString(value));
		}
		else
		{
			cell.setCellValue(new HSSFRichTextString(value));
		}
	}

	private boolean validateColIndex(int colIndex)
	{
		return colIndex > -1 && colIndex < activeWorkbook.getSpreadsheetVersion().getMaxColumns();
	}

	private boolean validateRowIndex(int rowIndex)
	{
		return rowIndex > -1 && rowIndex < activeWorkbook.getSpreadsheetVersion().getMaxRows();
	}

	private boolean isWorkbookPresent()
	{
		boolean result = true;
		if (Objects.isNull(activeWorkbook))
		{
			result = addErrorMessage("workbook missing (create workbook first)");
		}
		return result;
	}

//	private boolean isFunctionSupported(String function)
//	{
//		int pos = function.indexOf("(");
//		if (pos > -1)
//		{
//			String name = function.substring(0, pos - 1);
//			FunctionNameEval functionEval = new FunctionNameEval(name);
//			System.out.println(functionEval);
//		}
//		return true;
//	}

}