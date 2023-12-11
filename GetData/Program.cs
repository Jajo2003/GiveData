using System;
using System.Numerics;
using OfficeOpenXml;
using System.Net.Http;
using System.Collections.Generic;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using WebDriverManager.DriverConfigs.Impl;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Internal;
using System.Text.Json;
using System.Reflection.Metadata;


namespace retrievedata
{
	class dataClass
	{
		public static async Task Main()
		{

			Console.Write("Input your Path:");

			string pathtofile = Console.ReadLine();

			Console.Write("File name:");

			string filename = Console.ReadLine();

			string path = pathtofile + @"\" + filename + ".xlsx";

			List<string> Trackings = new List<string>();

			try
			{
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
				using (var package = new ExcelPackage(new FileInfo(path)))
				{

					var sheet = package.Workbook.Worksheets[0];


					int rowCount = sheet.Dimension.Rows;


					for (int i = 1; i <= rowCount; i++)
					{
						Trackings.Add(sheet.Cells["A" + $"{i}"].Text);
					}
					await SendData(Trackings);
				}

			}
			catch (Exception ex)
			{

				Console.WriteLine($"An error occurred: {ex.Message}");
			}
			finally
			{
				Console.WriteLine("\n\ndasrulebistvis daachiret sasurvel klavishs!!!");

				Console.ReadKey();
			}


		}

		


		public static async Task SendData(List<string> Trackings)
		{



			try
			{

				Console.Write("sheiyvanet wamebis raodenoba (wamebshi):");



				double seconds = Convert.ToDouble(Console.ReadLine());

				int miliseconds = (int)(seconds * 1000);

				Console.Write("Sheiyvanet Paroli:");

				string password = Console.ReadLine();

				string apiUrl = "https://www.moveasy.ge/api/rs-tracking";

				string apiKey = "FgnrgeDbESQrJPe8646po3sVDtYDZQEN";



				using (var driver = new ChromeDriver())
				{
					driver.Navigate().GoToUrl("https://decl.rs.ge/decls.aspx");

					driver.Manage().Window.Maximize();


					var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3000));

					IWebElement usernameField = driver.FindElement(By.Id("username"));
					IWebElement passwordField = driver.FindElement(By.Id("password"));
					IWebElement loginButton = driver.FindElement(By.Id("btnLogin"));

					await Task.Delay(2000);

					usernameField.SendKeys("404640411");
					passwordField.SendKeys(password);

					await Task.Delay(5000);


					loginButton.Click();

					await Task.Delay(15000);



					IWebElement OpenPage = driver.FindElement(By.ClassName("divModuleName"));
					await Task.Delay(5000);

					OpenPage.Click();

					await Task.Delay(5000);


					wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));



					await Task.Delay(5000);

					List<IWebElement> openModals = driver.FindElements(By.Id("control_0_smt")).ToList();

					wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

					await Task.Delay(5000);
					foreach (var modal in openModals)
					{

						if (modal.GetAttribute("innerHTML") == "<div>დარიდერება</div>")
						{
							modal.Click();
							break;
						}

					}
			
					await Task.Delay(3000);
					wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));


					IWebElement inputTag = driver.FindElement(By.ClassName("scan_postnumber"));
					await Task.Delay(3000);

					if (inputTag == null)
					{
						Console.WriteLine("tag not found");
					}
					else
					{
						wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

						IWebElement GetFullscreen = driver.FindElement(By.CssSelector(".window_panel .window_header .maximizeImg"));

						await Task.Delay(3000);

						GetFullscreen.Click();



						
						await Task.Delay(2000);
							foreach (var tracking in Trackings)
							{
								inputTag.Clear();
								await Task.Delay(miliseconds);

								inputTag.SendKeys(tracking);

								await Task.Delay(miliseconds);

								inputTag.SendKeys(Keys.Enter);

								await Task.Delay(1000);

								IWebElement ResultTag = driver.FindElement(By.Id("scan_result_parcel_status"));

								await Task.Delay(1000);

								string resultTagContent = ResultTag.GetAttribute("outerHTML");
							using (var httpClient = new HttpClient())
							{
								var url = $"{apiUrl}?key={apiKey}&tracking={tracking}&items={WebUtility.UrlEncode(resultTagContent)}";

								using (var request = new HttpRequestMessage(HttpMethod.Get, url))
								{
									var response = await httpClient.SendAsync(request);

									await Task.Delay(1000);
									if (response.IsSuccessStatusCode)
									{
										Console.WriteLine($"Tracking: {tracking} Gaigzavna ");
										
									}
									else
									{
										Console.WriteLine("Ver Gaigzavna");
									}
								}
							}




							await Task.Delay(1000);

								await Task.Delay(miliseconds);
							}
							await Task.Delay(5000);
						
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"{ex.Message} Error");
			}
		}


	}
	




}

