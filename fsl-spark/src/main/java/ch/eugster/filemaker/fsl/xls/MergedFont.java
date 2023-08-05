package ch.eugster.filemaker.fsl.xls;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

class MergedFont implements MessageProvider
{
	String name;
	Short size;
	Boolean bold;
	Boolean italic;
	Byte underline;
	Boolean strikeOut;
	Short typeOffset;
	Short color;

	public MergedFont(Font font)
	{
		name = font.getFontName();
		size = font.getFontHeightInPoints();
		bold = font.getBold();
		italic = font.getItalic();
		underline = font.getUnderline();
		strikeOut = font.getStrikeout();
		typeOffset = font.getTypeOffset();
		color = font.getColor();
	}

	public void applyToFont(Sheet sheet, Font font)
	{
		font.setFontName(name);
		font.setFontHeightInPoints(size);
		font.setBold(bold);
		font.setItalic(italic);
		font.setUnderline(underline);
		font.setStrikeout(strikeOut);
		font.setTypeOffset(typeOffset);
		font.setColor(color);
	}

	public boolean applyRequestedFontStyles(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		JsonNode nameNode = requestNode.findPath(Key.NAME.key());
		if (nameNode.isTextual())
		{
			name = nameNode.asText();
		}
		JsonNode sizeNode = requestNode.findPath(Key.SIZE.key());
		if (sizeNode.isInt())
		{
			int s = sizeNode.asInt();
			if (s >= 0 && s <= Short.MAX_VALUE)
			{
				size = (short) s;
			}
			else
			{
				result = addErrorMessage(responseNode, "invalid argument 'size' (" + s + ")");
			}
		}
		JsonNode styleNode = requestNode.findPath(Key.STYLE.key());
		if (styleNode.isInt())
		{
			switch ( styleNode.asInt())
			{
				case 1:
				{
					bold = true;
					italic = false;
					break;
				}
				case 2:
				{
					bold = false;
					italic = true;
					break;
				}
				case 3:
				{
					bold = true;
					italic = true;
					break;
				}
				default:
				{
					bold = false;
					italic = false;
					break;
				}
			}
		}
		else
		{
			JsonNode boldNode = requestNode.findPath(Key.BOLD.key());
			if (boldNode.isInt())
			{
				bold = boldNode.asBoolean();
			}
			JsonNode italicNode = requestNode.findPath(Key.ITALIC.key());
			if (italicNode.isInt())
			{
				italic = italicNode.asBoolean();
			}
		}
		JsonNode underlineNode = requestNode.findPath(Key.UNDERLINE.key());
		if (underlineNode.isTextual())
		{
			String u = underlineNode.asText();
			try
			{
				Underline ul = Underline.valueOf(u.toUpperCase());
				underline = (byte) ul.index;
			}
			catch (Exception e)
			{
				result = addErrorMessage(responseNode, "invalid argument 'underline' (" + u + ")");
			}
		}
		if (underlineNode.isInt())
		{
			int u = underlineNode.asInt();
			try
			{
				underline = (byte) Underline.index(u);
			}
			catch (Exception e)
			{
				result = addErrorMessage(responseNode, "invalid argument 'underline' (" + u + ")");
			}
		}
		JsonNode strikeOutNode = requestNode.findPath(Key.STRIKE_OUT.key());
		if (strikeOutNode.isInt())
		{
			strikeOut = strikeOutNode.asBoolean();
		}
		JsonNode typeOffsetNode = requestNode.findPath(Key.TYPE_OFFSET.key());
		if (typeOffsetNode.isTextual())
		{
			String t = typeOffsetNode.asText();
			try
			{
				TypeOffset to = TypeOffset.valueOf(t.toUpperCase());
				{
					typeOffset = (short) to.ordinal();
				}
			}
			catch (Exception e)
			{
				result = addErrorMessage(responseNode, "invalid argument 'type_offset' (" + t + ")");
			}
		}
		if (typeOffsetNode.isInt())
		{
			int t = typeOffsetNode.asInt();
			try
			{
				TypeOffset to = TypeOffset.values()[t];
				{
					typeOffset = (short) to.ordinal();
				}
			}
			catch (Exception e)
			{
				result = addErrorMessage(responseNode, "invalid argument 'type_offset' (" + t + ")");
			}
		}
		JsonNode colorNode = requestNode.findPath(Key.COLOR.key());
		if (!colorNode.isMissingNode())
		{
			if (colorNode.isTextual())
			{
				try
				{
					color = IndexedColors.valueOf(colorNode.asText().toUpperCase()).getIndex();
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "invalid argument 'foreground.color' (" + colorNode.asText() + ")");
				}
			}
			else if (colorNode.isInt())
			{
				try
				{
					color = IndexedColors.values()[colorNode.shortValue()].index;
				}
				catch (Exception e)
				{
					result = addErrorMessage(responseNode, "invalid argument 'foreground.color' (" + colorNode.asInt() + ")");
				}
			}
			else
			{
				result = addErrorMessage(responseNode, "invalid argument 'foreground.color' (" + colorNode.asText() + ")");
			}
		}
		return result;
	}

	public String getName()
	{
		return name;
	}

	public Short getSize()
	{
		return size;
	}

	public Boolean getBold()
	{
		return bold;
	}

	public Boolean getItalic()
	{
		return italic;
	}

	public Byte getUnderline()
	{
		return underline;
	}

	public Boolean getStrikeOut()
	{
		return strikeOut;
	}

	public Short getTypeOffset()
	{
		return typeOffset;
	}

	public Short getColor()
	{
		return color;
	}

	enum TypeOffset
	{
		SS_NORMAL, SS_SUPER, SS_SUB;
	}
	
	enum Underline
	{
		U_NONE(0),  U_SINGLE(1), U_DOUBLE(2), U_SINGLE_ACCOUNTING(33), U_DOUBLE_ACCOUNTING(34);

		private int index;
		
		private Underline(int index)
		{
			this.index = index;
		}
		
		public int index()
		{
			return this.index;
		}
		
		public static int index(int index) throws IllegalArgumentException
		{
			for (Underline u : Underline.values())
			{
				if (u.index() == index)
				{
					return index;
				}
			}
			throw new IllegalArgumentException();
		}
	}
}

