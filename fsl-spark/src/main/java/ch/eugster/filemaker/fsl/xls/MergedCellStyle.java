package ch.eugster.filemaker.fsl.xls;

import java.util.Objects;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.TextNode;

class MergedCellStyle implements MessageProvider
{
	private HorizontalAlignment halign;
	private VerticalAlignment valign;
	private BorderStyle bottom;
	private BorderStyle left;
	private BorderStyle right;
	private BorderStyle top;
	private Short bColor;
	private Short lColor;
	private Short rColor;
	private Short tColor;
	private Short dataFormat;
	private String format;
	private Short bgColor;
	private Short fgColor;
	private FillPatternType fillPattern;
	private Integer fontIndex;
	private Boolean shrinkToFit;
	private Boolean wrapText;

	public MergedCellStyle(CellStyle cellStyle)
	{
		halign = cellStyle.getAlignment();
		valign = cellStyle.getVerticalAlignment();
		bottom = cellStyle.getBorderBottom();
		left = cellStyle.getBorderLeft();
		right = cellStyle.getBorderRight();
		top = cellStyle.getBorderTop();
		bColor = cellStyle.getBottomBorderColor();
		lColor = cellStyle.getLeftBorderColor();
		rColor = cellStyle.getRightBorderColor();
		tColor = cellStyle.getTopBorderColor();
		dataFormat = cellStyle.getDataFormat();
		bgColor = cellStyle.getFillBackgroundColor();
		fgColor = cellStyle.getFillForegroundColor();
		fillPattern = cellStyle.getFillPattern();
		fontIndex = cellStyle.getFontIndex();
		shrinkToFit = cellStyle.getShrinkToFit();
		wrapText = cellStyle.getWrapText();
	}

	public void applyToCellStyle(Sheet sheet, CellStyle cellStyle)
	{
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(valign);
		cellStyle.setBorderBottom(bottom);
		cellStyle.setBorderLeft(left);
		cellStyle.setBorderRight(right);
		cellStyle.setBorderTop(top);
		cellStyle.setBottomBorderColor(bColor);
		cellStyle.setLeftBorderColor(lColor);
		cellStyle.setRightBorderColor(rColor);
		cellStyle.setTopBorderColor(tColor);
		if (dataFormat == -1)
		{
			DataFormat df = sheet.getWorkbook().createDataFormat();
			cellStyle.setDataFormat(df.getFormat(format));
		}
		else
		{
			cellStyle.setDataFormat(dataFormat);
		}
		cellStyle.setFillBackgroundColor(bgColor);
		cellStyle.setFillForegroundColor(fgColor);
		cellStyle.setFillPattern(fillPattern);
		cellStyle.setFont(sheet.getWorkbook().getFontAt(fontIndex));
		cellStyle.setShrinkToFit(shrinkToFit);
		cellStyle.setWrapText(wrapText);

	}

	public boolean applyRequestedStyles(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		JsonNode alignmentNode = requestNode.findPath(Key.ALIGNMENT.key());
		if (alignmentNode.isObject())
		{
			JsonNode horizontalNode = alignmentNode.findPath(Key.HORIZONTAL.key());
			if (horizontalNode.isTextual())
			{
				try
				{
					halign = HorizontalAlignment.valueOf(horizontalNode.asText().toUpperCase());
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'alignment.horizontal' (" + horizontalNode.asText() + ")");
				}
			}
			JsonNode verticalNode = alignmentNode.findPath(Key.VERTICAL.key());
			if (verticalNode.isTextual())
			{
				try
				{
					valign = VerticalAlignment.valueOf(verticalNode.asText().toUpperCase());
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'alignment.vertical' (" + verticalNode.asText() + ")");
				}
			}
		}
		JsonNode borderNode = requestNode.findPath(Key.BORDER.key());
		if (borderNode.isObject())
		{
			JsonNode styleNode = borderNode.findPath(Key.STYLE.key());
			if (styleNode.isObject())
			{
				JsonNode bottomNode = styleNode.findPath(Key.BOTTOM.key());
				if (bottomNode.isTextual())
				{
					BorderStyle borderStyle = valueOfBorderStyle(bottomNode.asText());
					if (Objects.nonNull(borderStyle))
					{
						bottom = borderStyle;
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.style.bottom' (" + bottomNode.asText() + ")");
					}
				}
				JsonNode leftNode = styleNode.findPath(Key.LEFT.key());
				if (leftNode.isTextual())
				{
					BorderStyle borderStyle = valueOfBorderStyle(leftNode.asText());
					if (Objects.nonNull(borderStyle))
					{
						left = borderStyle;
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.style.left' (" + leftNode.asText() + ")");
					}
				}
				JsonNode rightNode = styleNode.findPath(Key.RIGHT.key());
				if (rightNode.isTextual())
				{
					BorderStyle borderStyle = valueOfBorderStyle(rightNode.asText());
					if (Objects.nonNull(borderStyle))
					{
						right = borderStyle;
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.style.right' (" + rightNode.asText() + ")");
					}
				}
				JsonNode topNode = styleNode.findPath(Key.TOP.key());
				if (topNode.isTextual())
				{
					BorderStyle borderStyle = valueOfBorderStyle(topNode.asText());
					if (Objects.nonNull(borderStyle))
					{
						top = borderStyle;
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.style.top' (" + topNode.asText() + ")");
					}
				}
			}
			JsonNode colorNode = borderNode.findPath(Key.COLOR.key());
			if (!colorNode.isMissingNode())
			{
				JsonNode bottomNode = colorNode.findPath(Key.BOTTOM.key());
				if (!bottomNode.isMissingNode())
				{
					if (bottomNode.isTextual())
					{
						try
						{
							bColor = IndexedColors.valueOf(bottomNode.asText().toUpperCase()).getIndex();
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.bottom' (" + bottomNode.asText() + ")");
						}
					}
					else if (bottomNode.isInt())
					{
						try
						{
							bColor = IndexedColors.values()[bottomNode.shortValue()].index;
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.bottom' (" + bottomNode.asInt() + ")");
						}
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.color.bottom' (" + bottomNode.asText() + ")");
					}
				}
				JsonNode leftNode = colorNode.findPath(Key.LEFT.key());
				if (!leftNode.isMissingNode())
				{
					if (leftNode.isTextual())
					{
						try
						{
							lColor = IndexedColors.valueOf(leftNode.asText().toUpperCase()).getIndex();
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.left' (" + leftNode.asText() + ")");
						}
					}
					else if (leftNode.isInt())
					{
						try
						{
							lColor = IndexedColors.values()[leftNode.shortValue()].index;
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.left' (" + leftNode.asInt() + ")");
						}
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.color.left' (" + leftNode.asText() + ")");
					}
				}
				JsonNode rightNode = colorNode.findPath(Key.RIGHT.key());
				if (!rightNode.isMissingNode())
				{
					if (rightNode.isTextual())
					{
						try
						{
							rColor = IndexedColors.valueOf(rightNode.asText().toUpperCase()).getIndex();
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.right' (" + rightNode.asText() + ")");
						}
					}
					else if (rightNode.isInt())
					{
						try
						{
							rColor = IndexedColors.values()[rightNode.shortValue()].index;
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.right' (" + rightNode.asInt() + ")");
						}
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.color.right' (" + rightNode.asText() + ")");
					}
				}
				JsonNode topNode = colorNode.findPath(Key.TOP.key());
				if (!topNode.isMissingNode())
				{
					if (topNode.isTextual())
					{
						try
						{
							tColor = IndexedColors.valueOf(topNode.asText().toUpperCase()).getIndex();
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.top' (" + topNode.asText() + ")");
						}
					}
					else if (topNode.isInt())
					{
						try
						{
							tColor = IndexedColors.values()[topNode.shortValue()].index;
						}
						catch (Exception e)
						{
							result = addErrorMessage(responseNode, "illegal argument 'border.color.top' (" + topNode.asInt() + ")");
						}
					}
					else
					{
						result = addErrorMessage(responseNode, "illegal argument 'border.color.top' (" + topNode.asText() + ")");
					}
				}
			}
		}
		JsonNode dataFormatNode = requestNode.findPath(Key.DATA_FORMAT.key());
		if (dataFormatNode.isTextual())
		{
			format = dataFormatNode.asText();
			dataFormat = (short) BuiltinFormats.getBuiltinFormat(format);
		}
		else if (dataFormatNode.isInt())
		{
			dataFormat = (short) dataFormatNode.asInt();
		}
		JsonNode bgNode = requestNode.findPath(Key.BACKGROUND.key());
		if (bgNode.isObject())
		{
			JsonNode colorNode = bgNode.findPath(Key.COLOR.key());
			if (colorNode.isTextual())
			{
				try
				{
					bgColor = IndexedColors.valueOf(colorNode.asText().toUpperCase()).getIndex();
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'background.color' (" + colorNode.asText() + ")");
				}
			}
			else if (colorNode.isInt())
			{
				try
				{
					bgColor = IndexedColors.values()[colorNode.shortValue()].index;
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'background.color' (" + colorNode.asInt() + ")");
				}
			}
			else
			{
				result = addErrorMessage(responseNode, "illegal argument 'background.color' (" + colorNode.asText() + ")");
			}
		}
		JsonNode fgNode = requestNode.findPath(Key.FOREGROUND.key());
		if (fgNode.isObject())
		{
			JsonNode colorNode = fgNode.findPath(Key.COLOR.key());
			if (colorNode.isTextual())
			{
				try
				{
					fgColor = IndexedColors.valueOf(colorNode.asText().toUpperCase()).getIndex();
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'foreground.color' (" + colorNode.asText() + ")");
				}
			}
			else if (colorNode.isInt())
			{
				try
				{
					bgColor = IndexedColors.values()[colorNode.shortValue()].index;
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument 'foreground.color' (" + colorNode.asInt() + ")");
				}
			}
			else
			{
				result = addErrorMessage(responseNode, "illegal argument 'foreground.color' (" + colorNode.asText() + ")");
			}
		}
		JsonNode fillPatternNode = requestNode.findPath(Key.FILL_PATTERN.key());
		if (!fillPatternNode.isMissingNode())
		{
			if (TextNode.class.isInstance(fillPatternNode))
			{
				try
				{
					fillPattern = FillPatternType.valueOf(fillPatternNode.asText());
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "illegal argument '" + Key.FILL_PATTERN.key() + "' (" + fillPattern + ")");
				}
			}
		}
		JsonNode fontIndexNode = requestNode.findPath(Key.FONT.key());
		if (!fontIndexNode.isMissingNode())
		{
			if (fontIndexNode.isInt())
			{
				fontIndex = fontIndexNode.asInt();
			}
			else
			{
				result = addErrorMessage(responseNode, "illegal argument 'font' (" + fontIndex + ")");
			}
		}
		JsonNode shrinkToFitNode = requestNode.findPath(Key.SHRINK_TO_FIT.key());
		if (!shrinkToFitNode.isMissingNode())
		{
			if (shrinkToFitNode.isInt())
			{
				shrinkToFit = shrinkToFitNode.asBoolean();
			}
		}
		JsonNode wrapTextNode = requestNode.findPath(Key.WRAP_TEXT.key());
		if (!wrapTextNode.isMissingNode())
		{
			if (wrapTextNode.isInt())
			{
				wrapText = wrapTextNode.asBoolean();
			}
		}
		return result;
	}

	protected BorderStyle valueOfBorderStyle(String borderStyleName)
	{
		BorderStyle borderStyle = null;
		try
		{
			borderStyle = BorderStyle.valueOf(borderStyleName.toUpperCase());
		}
		catch (Exception e)
		{
		}
		return borderStyle;
	}

	public void setHalign(HorizontalAlignment halign)
	{
		this.halign = halign;
	}

	public void setValign(VerticalAlignment valign)
	{
		this.valign = valign;
	}

	public void setBottom(BorderStyle bottom)
	{
		this.bottom = bottom;
	}

	public void setLeft(BorderStyle left)
	{
		this.left = left;
	}

	public void setRight(BorderStyle right)
	{
		this.right = right;
	}

	public void setTop(BorderStyle top)
	{
		this.top = top;
	}

	public void setbColor(Short bColor)
	{
		this.bColor = bColor;
	}

	public void setlColor(Short lColor)
	{
		this.lColor = lColor;
	}

	public void setrColor(Short rColor)
	{
		this.rColor = rColor;
	}

	public void settColor(Short tColor)
	{
		this.tColor = tColor;
	}

	public void setDataFormat(Short dataFormat)
	{
		this.dataFormat = dataFormat;
	}

	public void setFormat(String format)
	{
		this.format = format;
	}

	public void setBgColor(Short bgColor)
	{
		this.bgColor = bgColor;
	}

	public void setFgColor(Short fgColor)
	{
		this.fgColor = fgColor;
	}

	public void setFillPattern(FillPatternType fillPattern)
	{
		this.fillPattern = fillPattern;
	}

	public void setFontIndex(Integer fontIndex)
	{
		this.fontIndex = fontIndex;
	}

	public void setShrinkToFit(Boolean shrinkToFit)
	{
		this.shrinkToFit = shrinkToFit;
	}

	public void setWrapText(Boolean wrapText)
	{
		this.wrapText = wrapText;
	}

	public HorizontalAlignment getHalign()
	{
		return halign;
	}

	public VerticalAlignment getValign()
	{
		return valign;
	}

	public BorderStyle getBottom()
	{
		return bottom;
	}

	public BorderStyle getLeft()
	{
		return left;
	}

	public BorderStyle getRight()
	{
		return right;
	}

	public BorderStyle getTop()
	{
		return top;
	}

	public Short getbColor()
	{
		return bColor;
	}

	public Short getlColor()
	{
		return lColor;
	}

	public Short getrColor()
	{
		return rColor;
	}

	public Short gettColor()
	{
		return tColor;
	}

	public Short getDataFormat()
	{
		return dataFormat;
	}

	public String getFormat()
	{
		return format;
	}

	public Short getBgColor()
	{
		return bgColor;
	}

	public Short getFgColor()
	{
		return fgColor;
	}

	public FillPatternType getFillPattern()
	{
		return fillPattern;
	}

	public Integer getFontIndex()
	{
		return fontIndex;
	}

	public Boolean getShrinkToFit()
	{
		return shrinkToFit;
	}

	public Boolean getWrapText()
	{
		return wrapText;
	}
}
