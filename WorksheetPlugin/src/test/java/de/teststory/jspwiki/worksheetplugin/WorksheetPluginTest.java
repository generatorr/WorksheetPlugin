package de.teststory.jspwiki.worksheetplugin;

import junit.framework.TestCase;

public class WorksheetPluginTest extends TestCase {

	public void testCleanCellValue() {

		// [0]=input [1]=expectedresult
		String[][] testMatrix = { //
				{ "abc", "abc" }, //0	
				{ "", " " }, //1
				{ null, " " }, //2
				{ " ", " " },//3
				{ "|", "-" }, //4
				{ "ab|cd|ef||", "ab-cd-ef--" }, //5
				{ "\r\n", "\\\\" }, //6
				{ "\r", "\r" }, //7
				{ "\n", "\\\\" }, //8
				{ "\n\r\n", "\\\\\\\\" }, //9
				{ "\nab\r\n|cd||\n\r", "\\\\ab\\\\-cd--\\\\\r" }, //10
		};
		for (int i = 0, size = testMatrix.length; i < size; ++i) {
			String input = testMatrix[i][0];
			String expectedResult = testMatrix[i][1];
			assertEquals(String.valueOf(i), expectedResult, WorksheetPlugin.cleanCellValue(input));
		}

	}

}
