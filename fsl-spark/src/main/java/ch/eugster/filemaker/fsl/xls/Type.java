package ch.eugster.filemaker.fsl.xls;

public enum Type
{
	XLSX("xlsx"), XLS("xls");

	private String extension;
	
	private Type(String extension)
	{
		this.extension = extension;
	}
	
	public String extension()
	{
		return this.extension;
	}
	
	public static Type findByExtension(String extension)
	{
		for (Type type : Type.values())
		{
			if (type.extension.equals(extension))
			{
				return type;
			}
		}
		return null;
	}
}
