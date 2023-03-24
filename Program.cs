//C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
using System;
using System.IO;
using System.Data.SqlClient;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace TFlexRepositoryCleaner
{

public class Macro
{
    static void Main()
    { 
			//параметры приложения из ini файла
			string dirPath = "";
			string lastDate = "";
			bool imitation = true;
			bool filledComment = false;
			string dataSource = "";
			string userID = "";
			string password = "";
			string initialCatalog = "";
			try
			{
				//запрашиваем параметры из ini файла
				IniFile ini = new IniFile("conf.ini");
			if (ini.KeyExists("General", "dirPath"))
				dirPath = ini.ReadINI("General", "dirPath");
			if (ini.KeyExists("General", "lastDate"))
				lastDate = ini.ReadINI("General", "lastDate");
			if (ini.KeyExists("General", "imitation"))
				imitation = Convert.ToBoolean(ini.ReadINI("General", "imitation"));		
			if (ini.KeyExists("General", "filledComment"))		
				filledComment = Convert.ToBoolean(ini.ReadINI("General", "filledComment"));
			if (ini.KeyExists("connection", "dataSource"))
				dataSource = ini.ReadINI("connection", "dataSource");
			if (ini.KeyExists("connection", "userID"))
				userID = ini.ReadINI("connection", "userID");
			if (ini.KeyExists("connection", "password"))
				password = ini.ReadINI("connection", "password");
			if (ini.KeyExists("connection", "initialCatalog"))
				initialCatalog = ini.ReadINI("connection", "initialCatalog");
			//создаем экземпляр объекта с параметрами
			TFlexRepoCleaner TFRC = new TFlexRepoCleaner(dirPath, lastDate, imitation, filledComment, 
															dataSource, userID, password, initialCatalog);
			TFRC.CleanRepo();													
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}  
    }

}

class TFlexRepoCleaner
{
	private string _dirPath {get; set;}
  private string _lastDate {get; set;}
	private bool _imitation {get; set;}
	private bool _filledComment {get; set;}
	private string _dataSource {get; set;} 
  private string _userID {get; set;}            
  private string _password {get; set;}    
  private string _initialCatalog {get; set;}
  public TFlexRepoCleaner (string dirPath, string lastDate, bool imitation, bool filledComment, 
													string dataSource, string userID, string password, string initialCatalog)
	{
		_dirPath = dirPath;
		_lastDate = lastDate;
		_imitation = imitation;
		_filledComment = filledComment;
		_dataSource = dataSource;
		_userID = userID;
		_password = password;
		_initialCatalog = initialCatalog;
	}

	public void CleanRepo()
	{
		string filePath = ""; 
		string filledCommentCond = "";
		int objCount = 0;
		if (!_filledComment)
		filledCommentCond = "c.Comment = '' and ";  
		StringBuilder sb = new StringBuilder();
		sb.Append("<html> <head> <style> table {border-collapse: collapse; width: 90%;} 	table, td { 	} 	td,th { 	  padding: 5px; 	  text-align: left; 	  border-bottom: 1px solid #ddd; 	} 	.stripe:nth-child(even) { 	  background-color: #f2f2f2; 	} .red {background-color: #E53935; }  </style> </head> <body> <table> <tr>   <th>ID</th>   <th>Имя файла</th>   <th>Версия</th> <th>Дата последнего изменения</th> </tr>");
		try 
    { 	  
       SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
       builder.DataSource = _dataSource; 
       builder.UserID = _userID;            
       builder.Password = _password;     
       builder.InitialCatalog = _initialCatalog;
			 using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
       {
          connection.Open(); 
          String sql = "SELECT co.ObjectID, co.ObjectVersion, f.Name, f.Path, f.s_EditDate " +
												"FROM [TFlexDOCs].[dbo].[ChangelistObjects] co " +
												"INNER JOIN [TFlexDOCs].[dbo].[Files] f ON co.ObjectID = f.s_ObjectID AND co.ObjectVersion = f.s_Version " +
												"INNER JOIN [TFlexDOCs].[dbo].[Changelists] c ON co.ChangelistID = c.PK " +
												"WHERE ("+filledCommentCond+" c.Label = '' and f.s_ActualVersion = 0 and f.s_EditDate<'"+_lastDate+"' and f.s_ClassID <> 55) "+
												"ORDER BY co.ObjectID;";
					List<FileParams> fileList = new List<FileParams>();					
					//Выборка неактуальных и неименованных версий
          using (SqlCommand command = new SqlCommand(sql, connection))
          {
            using (SqlDataReader reader = command.ExecuteReader())
            {
              while (reader.Read())
              {
								objCount++;
								FileParams f = new FileParams(); 
								f.dbFileID = reader.GetInt32(0).ToString();
								f.dbFileVer = reader.GetInt32(1).ToString();
								f.dbFileName = reader.GetString(2);
								f.dbFilePath = reader.GetString(3);
								f.dbFileEditDate = reader.GetDateTime(4);
								fileList.Add(f);
              }
						}						 								 		
          }
					//Удаление файлов и записей о версиях из БД, если выключен режим имитации
					foreach (FileParams fi in fileList)
					{
						filePath = _dirPath + fi.dbFileID +"\\" + fi.dbFileVer;
						FileInfo file = new FileInfo(filePath);
						if (!_imitation)
						{
							if (File.Exists(filePath))
							{
							  try 
							  {
								  file.Delete();
								  string sqlDel1 = "DELETE FROM [TFlexDOCs].[dbo].[ChangelistObjects] "+
															"WHERE ObjectID = "+fi.dbFileID+" AND ObjectVersion = "+fi.dbFileVer+";";
								  string sqlDel2 = "DELETE FROM [TFlexDOCs].[dbo].[Files] "+
															"WHERE s_ObjectID = "+fi.dbFileID+" AND s_Version = "+fi.dbFileVer+";";
								  using (SqlCommand delCommand1 = new SqlCommand(sqlDel1, connection))
								  {
									  int rowCount = delCommand1.ExecuteNonQuery();
								  }	
								  using (SqlCommand delCommand1 = new SqlCommand(sqlDel2, connection))
								  {
									  int rowCount = delCommand1.ExecuteNonQuery();
								  }
									sb.Append(printReport(fi.dbFileID, fi.dbFileVer, fi.dbFileName, fi.dbFilePath, fi.dbFileEditDate, false));
							  }	
							  catch
							  {
									Console.WriteLine("1");
								  sb.Append(printReport(fi.dbFileID, fi.dbFileVer, fi.dbFileName, fi.dbFilePath, fi.dbFileEditDate, true));
							  }
							}
							else 
							{
								Console.WriteLine(filePath);
								sb.Append(printReport(fi.dbFileID, fi.dbFileVer, fi.dbFileName, fi.dbFilePath, fi.dbFileEditDate, true));
							}
						}
						else
						{
							sb.Append(printReport(fi.dbFileID, fi.dbFileVer, fi.dbFileName, fi.dbFilePath, fi.dbFileEditDate, false));
						}
					}	
					sb.Append("</table><h3>Всего объектов: "+objCount+"</h3> </body> </html>");
					using (StreamWriter outputFile = new StreamWriter(@"report.html")) 
					{
        		outputFile.WriteLine(sb.ToString());  
        	}							
		}                    
    }
    catch (SqlException e)
    {
      Console.WriteLine(e.ToString());
    }
	}         
	//формирование отчета
	private string printReport(string fileID, string fileVer, string fileName, string filePath, DateTime fileEditDate, bool error)
		{
			StringBuilder sb = new StringBuilder();
			if (error)
				sb.Append("<tr class='red'>");
			else
				sb.Append("<tr class='stripe'>");
			sb.Append("<td>"+fileID+"</td>");
			sb.Append("<td>"+filePath+ "\\" + fileName + "</td>");
			sb.Append("<td>"+fileVer+"</td>");
			sb.Append("<td>"+fileEditDate.ToString()+"</td>");
			sb.Append("</tr>");
			return sb.ToString();
		}
	//структура, хранящая данные о текущем файле для удаления	
	private struct FileParams
	{
		public string dbFileID;
		public string dbFileVer;
		public string dbFileName;
		public string dbFilePath;
		public DateTime dbFileEditDate;
	}	
}
//класс для работы с файлом конфигурации
class IniFile
{
  string _path; //Имя файла.
	[DllImport("kernel32")] // Подключаем kernel32.dll и описываем его функцию WritePrivateProfilesString
  static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);
  [DllImport("kernel32")] // Еще раз подключаем kernel32.dll, а теперь описываем функцию GetPrivateProfileString
  static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);
	public IniFile(string path)
  {
    _path = new FileInfo(path).FullName.ToString();
  }
	//Читаем ini-файл и возвращаем значение указного ключа из заданной секции.
  public string ReadINI(string Section, string Key)
  {
    var RetVal = new StringBuilder(255);
    GetPrivateProfileString(Section, Key, "", RetVal, 255, _path);
    return RetVal.ToString();
  }

	//Проверяем, есть ли такой ключ, в этой секции
  public bool KeyExists(string Section, string Key)
  {
    return ReadINI(Section, Key).Length > 0;
  }
} 
}
