using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace UnpackBooks
{
	class Program
	{
		static void Main(string[] args)
		{
			foreach(string file in Directory.EnumerateFiles(".","*.rar", SearchOption.TopDirectoryOnly))
			{
				string FileNameWithoutExt = Path.GetFileNameWithoutExtension(file);

				//解压 到 以名称命名的目录
				UnRarFile(file, FileNameWithoutExt);

				//移动文件
				MoveFile(FileNameWithoutExt);
			}

			Console.WriteLine("任意键退出!");
			Console.ReadKey(false);
		}

		static void MoveFile(string FileNameWithoutExt)
		{
			//过滤得到txt文件
			string TargetFile = Directory.EnumerateFiles(FileNameWithoutExt, "*.txt", SearchOption.TopDirectoryOnly).Where(x => !x.Contains("说明")).FirstOrDefault();

			//文件的新路径
			string NewFile = FileNameWithoutExt + Path.GetExtension(TargetFile);

			//新文件存在则删除
			if (File.Exists(NewFile)) File.Delete(NewFile);

			Console.WriteLine("{0} --> {1}", TargetFile, FileNameWithoutExt + Path.GetExtension(TargetFile));

			//移动到本目录
			File.Move(TargetFile, FileNameWithoutExt + Path.GetExtension(TargetFile));

			Directory.Delete(FileNameWithoutExt, true);
		}

		static void UnRarFile(string filepath, string dest)
		{
			if (Directory.Exists(dest)) Directory.Delete(dest, true);

			string shellArguments = string.Format(@"X {0} *.txt {1}\", Path.GetFileName(filepath), dest);

			//用Process调用
			using (Process unrar = new Process())
			{
				unrar.StartInfo.FileName = @"C:\Program Files\WinRAR\WinRAR.exe";
				unrar.StartInfo.Arguments = shellArguments;
				//隐藏rar本身的窗口
				//unrar.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
				unrar.Start();
				//等待解压完成
				unrar.WaitForExit();
				unrar.Close();
			}
		}

	}
}
