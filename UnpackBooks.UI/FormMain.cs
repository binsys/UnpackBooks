using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace UnpackBooks.UI
{
	public partial class FormMain : Form
	{
		public FormMain()
		{
			InitializeComponent();
		}

		private void FormMain_DragDrop(object sender, DragEventArgs e)
		{
			string SourcePath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

			if (System.IO.Directory.Exists(SourcePath))
			{
				foreach (string file in Directory.EnumerateFiles(SourcePath, "*.rar", SearchOption.TopDirectoryOnly))
				{
					// file = "D:\User\desktop\12.15\传奇纨绔少爷.rar"
					string FileNameWithoutExt = Path.GetFileNameWithoutExtension(file);

					//解压 到 以名称命名的目录
					UnRarFile(file, Path.Combine(SourcePath, FileNameWithoutExt));

					//移动文件
					MoveFile(FileNameWithoutExt, SourcePath);
				}

				MessageBox.Show("完毕！");
			}

		}

		private void FormMain_DragEnter(object sender, DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
				e.Effect = DragDropEffects.Link;
			else e.Effect = DragDropEffects.None;

		}


		static void MoveFile(string FileNameWithoutExt, string SourcePath)
		{
			//过滤得到txt文件
			string TargetFile = Directory.EnumerateFiles( Path.Combine(SourcePath, FileNameWithoutExt), "*.txt", SearchOption.TopDirectoryOnly).Where(x => !x.Contains("说明")).FirstOrDefault();

			//文件的新路径
			string NewFile = Path.Combine(SourcePath, FileNameWithoutExt + Path.GetExtension(TargetFile));

			//新文件存在则删除
			if (File.Exists(NewFile)) File.Delete(NewFile);

			//移动到 SourcePath 目录
			File.Move(TargetFile, NewFile);

			Directory.Delete(Path.Combine(SourcePath, FileNameWithoutExt), true);
		}

		static void UnRarFile(string filepath, string dest)
		{
			if (Directory.Exists(dest)) Directory.Delete(dest, true);

			string shellArguments = string.Format(@"X {0} *.txt {1}\",filepath, dest);

			//用Process调用
			using (Process unrar = new Process())
			{
				unrar.StartInfo.FileName = @"C:\Program Files\WinRAR\unrar.exe";
				unrar.StartInfo.Arguments = shellArguments;
				//隐藏rar本身的窗口
				unrar.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
				unrar.Start();
				//等待解压完成
				unrar.WaitForExit();
				unrar.Close();
			}
		}
	}
}
