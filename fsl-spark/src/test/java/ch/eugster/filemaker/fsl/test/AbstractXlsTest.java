package ch.eugster.filemaker.fsl.test;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;

import com.fasterxml.jackson.databind.ObjectMapper;

import ch.eugster.filemaker.fsl.xls.Xls;

public abstract class AbstractXlsTest extends AbstractTest
{
	protected ObjectMapper mapper = new ObjectMapper();
	
	protected final String WORKBOOK_1 = "./src/test/results/workbook1.xlsx";
	
	protected final String SHEET0 = "Sheet0";

	@BeforeEach
	protected void beforeEach() throws Exception
	{
	}
	
	@AfterEach
	protected void afterEach() throws Exception
	{
		Xls.activeWorkbook = null;
	}
}
