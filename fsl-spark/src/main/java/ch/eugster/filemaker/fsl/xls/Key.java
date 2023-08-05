package ch.eugster.filemaker.fsl.xls;

public enum Key
{
	// @formatter:off
	ALIGNMENT("alignment"),
	BACKGROUND("background"),
	BOLD("bold"),
	BORDER("border"),
	BOTTOM_RIGHT("bottom_right"),
	BOTTOM("bottom"),
	CELL("cell"),
	CENTER("center"),
	COL("col"),
	COLOR("color"),
	COPIES("copies"),
	DATA_FORMAT("data_format"),
	DIRECTION("direction"),
	FILL_PATTERN("fill_pattern"),
	FONT("font"),
	FOREGROUND("foreground"),
	FORMAT("format"),
	HORIZONTAL("horizontal"),
	INDEX("index"),
	ITALIC("italic"),
	LEFT("left"),
	NAME("name"),
	ORIENTATION("orientation"),
	PATH("path"),
	RANGE("range"),
	RIGHT("right"),
	ROTATION("rotation"),
	ROW("row"),
	SHEET("sheet"),
	SHRINK_TO_FIT("shrink_to_fit"),
	SOURCE("source"),
	SIZE("size"),
	STRIKE_OUT("strike_out"),
	STYLE("style"),
	TARGET("target"),
	TOP("top"),
	TOP_LEFT("top_left"),
	TYPE("type"),
	TYPE_OFFSET("type_offset"),
	UNDERLINE("underline"),
	VALUES("values"),
	VERTICAL("vertical"),
	WORKBOOK("workbook"),
	WRAP_TEXT("wrap_text");
	// @formatter:on

	private Key(String key)
	{
		this.key = key;
	}

	private String key;

	public String key()
	{
		return key;
	}
}

