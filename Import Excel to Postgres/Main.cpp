#include <iostream>
#include <xlnt/xlnt.hpp>
#include <nanodbc/nanodbc.h>
#include <vector>
#include "Converter.h"
#include <fstream>
#include <filesystem>
#include <thread>

typedef struct Settings
{
    std::vector <std::string> columnNames;
    std::string odbcname = "dsn=Wanted";
    std::string folder = "C:\\Software\\xls";
    // папка, куда где будут проводиться все махинации
    std::string workdir = "C:\\Software\\temp";
    std::string tableName = "wanted";
    bool deletexlsfile = true;
    bool deletedebugfiles = true;
    std::string filename = "";
    std::string logFileName = "Log.txt";
    std::string settingsfile = "settings.ini";
    std::string minute = "1";
};

std::string currentTime();
bool iniParse(Settings& params);
std::string constructInsertScript(nanodbc::result& results, Settings& setting);
std::vector<std::filesystem::path> findFile(std::string path);

int main()
{
    setlocale(LC_ALL, "Russian");
    Settings setting;

    std::ofstream logfile(setting.logFileName, std::ios::app);

	//проверяем ини файл, если нет то закрываемся
    if (!iniParse(setting))
    {
        logfile << currentTime() << "Can't open settings.ini" << std::endl;
        return -1;
    }

	logfile << currentTime() << "Starting..." << std::endl;


	bool waiting = true;

	while (true)
	{
		//"обнуляем" путь до файла, что бы не заходить в БД и ничего там не стереть
		setting.filename = "";

		//если папки к файлу нет, то пишем в лог и выключаемся
		if (!std::filesystem::exists(setting.folder))
		{
			logfile << currentTime() << "Path: " << setting.folder << " - not exist" << std::endl;
			logfile << currentTime() << "Close app with error..." << std::endl;
			return -1;
		}

		std::vector<std::filesystem::path> array;
		array = findFile(setting.folder);		

		// конвертируем полученный вектор в формате std::filesystem::path в string и записываем путь до файла в структуру
		setting.filename = array[0].string();

		//переименовываем файл, иначе если русские буквы в имени - то открыть не можем.....
		const char* oldName = setting.filename.c_str();
		std::string tempPath = setting.folder + "\\" + "1.xlsx";
		const char* newName = tempPath.c_str();

		rename(oldName, newName);

		array.clear();
		array = findFile(setting.folder);

		// конвертируем полученный вектор в формате std::filesystem::path в string и записываем путь до файла в структуру
		setting.filename = array[0].string();

		if (setting.filename != "")
		{	
			//создаем скрипт на очистку таблицы
			std::string clearTableScript = "delete from public." + setting.tableName;
			//создаем соединение с postgres через odbc x64
			nanodbc::connection connection(setting.odbcname);

			//создаем переменную класса workbook и загружаем в нее наш эксель файл
			xlnt::workbook wb;
			wb.load(setting.filename);
			auto ws = wb.active_sheet();

			//считаем сколько колонок в файле, что бы потом сравнить с таблицей
			short cells = 0;
			for (auto row : ws.rows(false))
			{
				for (auto cell : row)
				{
					cells++;
				}
				break;
			}

			//создаем переменную и записываем в нее результат скрипта, что бы понять сколько колонок у нас есть
			nanodbc::result results;
			results = nanodbc::execute(connection, "select * from public.wanted limit 1;");

			//если колличество столбцов в файле совпало с количеством столбцов в таблице, то работаем, если нет то пишем гадости в лог и закрываемся
			if (cells == results.columns())
			{
				//чистим таблицу
				nanodbc::execute(connection, clearTableScript);



				
				for (auto row : ws.rows(false))
				{
					std::string insertTableScript = "insert into public." + setting.tableName + "(";

					//добавляем к insertTableScript названия столбцов
					for (short i = 0; i < results.columns(); i++)
					{
						if (i + 1 < results.columns())
						{
							insertTableScript += results.column_name(i) + ", ";
						}
						else
						{
							insertTableScript += results.column_name(i) + ") values(";
						}
					}

					int i = 0;
					int count = 0;
					for (auto cell : row)
					{
						std::string value = "";
						value = UTF8ToANSI(cell.to_string());
						if (value == "")
						{
							count++;
						}

						if (i + 1 < results.columns())
						{
							insertTableScript += "\'" + value + "\', ";

						}
						else
						{
							insertTableScript += "\'" + value + "\'" + "); ";
						}
						i++;
					}
					if (count != 2)
					{
						nanodbc::execute(connection, insertTableScript);
					}	
				}

				logfile << currentTime() << "Successfully imported" << std::endl;
				remove(setting.filename.c_str());
			}
			else
			{
				logfile << currentTime() << "Incorrect file! the number of columns in the file does not correspond to the number of columns in the table" << std::endl;
				logfile << currentTime() << "Close app with error..." << std::endl;
				return -1;
			}
		}

		if(waiting)
			logfile << currentTime() << "Waiting..." << std::endl;

		std::chrono::milliseconds timespan(stoi(setting.minute) * 1000 * 60);
		std::this_thread::sleep_for(timespan);
	}

}





std::string currentTime()
{
    std::time_t time = std::time(0);
    std::tm* now = std::localtime(&time);
    char buffer[80];
    strftime(buffer, 80, "%d.%m.%Y %H:%M%p ", now);

    return buffer;
}
bool iniParse(Settings& params)
{
	std::ifstream iniFile(params.settingsfile);

	if (!iniFile.is_open())
	{
		return false;
	}

	std::string str;
	while (std::getline(iniFile, str))
	{
		std::string temp = "";
		for (int i = 0; i < str.length(); i++)
		{
			temp += str[i];
			if (temp == "odbcname")
			{
				params.odbcname = "dsn=";
				for (int j = i + 2; j < str.length(); j++)
				{
					params.odbcname += str[j];
				}
			}
			else if (temp == "tablename")
			{
				params.tableName = "";
				for (int j = i + 2; j < str.length(); j++)
				{
					params.tableName += str[j];
				}
			}
			else if (temp == "folder")
			{
				params.folder = "";
				for (int j = i + 2; j < str.length(); j++)
				{
					params.folder += str[j];
				}
			}
			else if (temp == "workdir")
			{
				params.workdir = "";
				for (int j = i + 2; j < str.length(); j++)
				{
					params.workdir += str[j];
				}
			}
			else if (temp == "deletexlsfile")
			{
				params.deletexlsfile = "";
				std::string deletebool = "";
				for (int j = i + 2; j < str.length(); j++)
				{

					deletebool += str[j];
					if (deletebool == "true")
					{
						params.deletexlsfile = true;
					}
					else if (deletebool == "false")
					{
						params.deletexlsfile = false;
					}
				}
			}
			else if (temp == "deletedebugfiles")
			{
				params.deletedebugfiles = "";
				std::string deletebool = "";
				for (int j = i + 2; j < str.length(); j++)
				{

					deletebool += str[j];
					if (deletebool == "true")
					{
						params.deletedebugfiles = true;
					}
					else if (deletebool == "false")
					{
						params.deletedebugfiles = false;
					}
				}
			}
			else if (temp == "minute")
			{
				params.minute = "";
				for (int j = i + 2; j < str.length(); j++)
				{
					params.minute += str[j];
				}
			}
		}
	}

	return true;
}

std::string constructInsertScript(nanodbc::result& results, Settings& setting)
{
	std::string sql = "insert into public." + setting.tableName + "(";

	const short columns = results.columns();

	for (short i = 0; i < columns; i++)
	{
		sql += results.column_name(i) + " character varying";
		if (i + 1 < columns)
		{
			sql += ", ";
		}
		else
			sql += ");";
		//std::cout << results.column_name(i);
	}

	return sql;
}


//создаем вектор<std::filesystem::path>, записываем в него название появившегося файла в папке 
std::vector<std::filesystem::path> findFile(std::string path)
{
	auto it = std::filesystem::directory_iterator(path);
	std::vector<std::filesystem::path> array;
	std::copy_if(std::filesystem::begin(it), std::filesystem::end(it), std::back_inserter(array),
		[](const auto& entry) {
			return std::filesystem::is_regular_file(entry);
		});

	return array;
}