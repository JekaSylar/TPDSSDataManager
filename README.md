# TPDSS Data Manager

TPDSS Data Manager is a Windows desktop application built with WPF and C#. It is specifically designed to automate the processing, splitting, and merging of Microsoft Access databases (`.accdb`) for regional statistical departments. 

## 🚀 Features

* **Database Splitting:** Analyzes a master Access database, builds a hierarchical tree of structural units, and allows users to extract specific departments into separate `.accdb` files.
* **Database Merging:** Merges multiple separated Access databases back into a single consolidated master file without duplicating records.
* **Asynchronous Processing:** Utilizes background tasks (`Task.Run`) and modern `IProgress<T>` interfaces to keep the UI responsive during heavy database operations.
* **Modern UI:** Clean, intuitive WPF interface with progress tracking and real-time status updates.

## 🛠 Prerequisites & Requirements

To run or build this project, your system must meet the following requirements:

1. **.NET 10.0** Ensure you have the .NET 10.0 SDK or Runtime installed on your machine.
2. **Microsoft Access Database Engine (DAO)**
   The application interacts with `.accdb` files using `Microsoft.Office.Interop.Access.Dao`. You must have one of the following installed:
   * Microsoft Office (with MS Access included) installed on your system.
   * **OR** the standalone [Microsoft Access Database Engine Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (Make sure the architecture, x86 or x64, matches your application's build target).

## 💻 Tech Stack

* **Language:** C# 
* **Framework:** .NET 10.0
* **UI:** Windows Presentation Foundation (WPF)
* **Database Interop:** DAO (Data Access Objects) via Office Interop

