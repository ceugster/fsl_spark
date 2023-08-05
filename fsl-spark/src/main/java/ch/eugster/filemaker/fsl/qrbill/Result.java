package ch.eugster.filemaker.fsl.qrbill;

import java.io.File;

public class Result
{
	private boolean result;

	private File file;

	public void setResult(boolean result)
	{
		this.result = result;
	}

	public boolean getResult()
	{
		return this.result;
	}

	public void setFile(File file)
	{
		this.file = file;
	}

	public File getFile()
	{
		return this.file;
	}
}
