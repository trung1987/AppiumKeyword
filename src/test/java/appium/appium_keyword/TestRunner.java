package appium.appium_keyword;

import java.io.IOException;

import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import appium.Utils.Constant;
import appium.Utils.ReadExcel;
import io.appium.java_client.android.AndroidDriver;

public class TestRunner {
	AndroidDriver<WebElement> androiddriver;
	@Test
	public void TestApp() throws RuntimeException, Exception{
		androiddriver=ReadExcel.ReadAll(Constant.File_name);
		
	}
	
}
