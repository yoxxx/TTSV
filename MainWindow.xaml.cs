/*
 * MainWindow.xaml.cs
 * Copyright c 2018 yo_xxx
 *
 * This file is part of TTSV.
 *
 * TTSV is free software; you can redistribute it and/or
 * modify it under the terms of the BSD License.
 *
 * TTSV is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;

namespace TTSV
{
    /// <summary>
    /// 各種動作設定を扱うクラス
    /// </summary>
    public class TTSVConfig
    {
        public string compID;
        public string vVoiceName;
        public string Comment;
        public string OpenJTalkOption;
        public string Prefix;
        public Double BaseNoteNum;
        public Double MiddleNoteNum;
        public Double TopNoteNum;
        public Double NaturalEndNoteNum;
        public Double BottomNoteNum;
        public Double QuestionNoteNum;
        public Double NoteSplitMode;
        public byte ColorR;
        public byte ColorG;
        public byte ColorB;

        /// <summary>
        /// 各種動作設定を扱うクラスのコンストラクタ
        /// </summary>
        /// <param name="sConfigpass">設定ファイルのパス</param>
        public TTSVConfig(string sConfigpass)
        {

            if (!System.IO.File.Exists(sConfigpass))
            {
                return;
            };

            System.IO.StreamReader srConfig = new System.IO.StreamReader(sConfigpass, Encoding.GetEncoding("SHIFT_JIS"));

            string CFG_Line = srConfig.ReadLine();
            string[] item;

            while (srConfig.EndOfStream == false)
            {
                item = CFG_Line.Split('=');
                item[0] = item[0].Trim();
                item[1] = item[1].Trim();
                switch (item[0])
                {
                    case "compID":
                        compID = item[1];
                        break;
                    case "vVoiceName":
                        vVoiceName = item[1];
                        break;
                    case "Comment":
                        Comment = item[1];
                        break;
                    case "OpenJTalkOption":
                        OpenJTalkOption = item[1];
                        break;
                    case "Prefix":
                        Prefix = item[1];
                        break;
                    case "BaseNoteNum":
                        BaseNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= BaseNoteNum) & (BaseNoteNum <= 127))
                        {
                            BaseNoteNum = BaseNoteNum * 1.00;
                        }
                        else
                        {
                            BaseNoteNum = 62.00;
                        }
                        break;
                    case "MiddleNoteNum":
                        MiddleNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= MiddleNoteNum) & (MiddleNoteNum <= 127))
                        {
                            MiddleNoteNum = MiddleNoteNum * 1.00;
                        }
                        else
                        {
                            MiddleNoteNum = 67.00;
                        }
                        break;
                    case "TopNoteNum":
                        TopNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= TopNoteNum) & (TopNoteNum <= 127))
                        {
                            TopNoteNum = TopNoteNum * 1.00;
                        }
                        else
                        {
                            TopNoteNum = 70.00;
                        }
                        break;
                    case "NaturalEndNoteNum":
                        NaturalEndNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= NaturalEndNoteNum) & (NaturalEndNoteNum <= 127))
                        {
                            NaturalEndNoteNum = NaturalEndNoteNum * 1.00;
                        }
                        else
                        {
                            NaturalEndNoteNum = 68.00;
                        }
                        break;
                    case "BottomNoteNum":
                        BottomNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= BottomNoteNum) & (BottomNoteNum <= 127))
                        {
                            BottomNoteNum = BottomNoteNum * 1.00;
                        }
                        else
                        {
                            BottomNoteNum = 60.00;
                        }
                        break;
                    case "QuestionNoteNum":
                        QuestionNoteNum = Convert.ToInt32(item[1]);
                        if ((0 <= QuestionNoteNum) & (QuestionNoteNum <= 127))
                        {
                            QuestionNoteNum = QuestionNoteNum * 1.00;
                        }
                        else
                        {
                            QuestionNoteNum = 64.00;
                        }
                        break;
                    case "NoteSplitMode":
                        NoteSplitMode = Convert.ToInt32(item[1]);
                        if ((1 <= NoteSplitMode) & (NoteSplitMode <= 2))
                        {
                            NoteSplitMode = NoteSplitMode * 1.00;
                        }
                        else
                        {
                            NoteSplitMode = 1.00;
                        }
                        break;
                    case "ColorR":
                        ColorR = Convert.ToByte(item[1]);
                        if (((byte)0 <= ColorR) & (ColorR <= (byte)255))
                        {
                            ColorR = (byte)(ColorR + 0);
                        }
                        else
                        {
                            ColorR = 255;
                        }
                        break;
                    case "ColorG":
                        ColorG = Convert.ToByte(item[1]);
                        if (((byte)0 <= ColorG) & (ColorG <= (byte)255))
                        {
                            ColorG = (byte)(ColorG + 0);
                        }
                        else
                        {
                            ColorG = 255;
                        }
                        break;
                    case "ColorB":
                        ColorB = Convert.ToByte(item[1]);
                        if (((byte)0 <= ColorB) & (ColorB <= (byte)255))
                        {
                            ColorB = (byte)(ColorB + 0);
                        }
                        else
                        {
                            ColorB = 255;
                        }
                        break;
                    default:
                        break;
                }
                CFG_Line = srConfig.ReadLine();
            }
        }
    }

    /// <summary>
    /// メインウィンドウのMainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {

        /// <summary>
        /// メインウィンドウ初期設定
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            //画面項目初期設定
            EntryText.Text = "";
            ConfigFilePass.Content = System.AppDomain.CurrentDomain.BaseDirectory + "TTSVConfig.txt";
            EntryTextPass.Content  = System.AppDomain.CurrentDomain.BaseDirectory + "ENTRY"; 
            VSQXPass.Content       = System.AppDomain.CurrentDomain.BaseDirectory + "VSQX";
            WorkFilePass.Content   = System.AppDomain.CurrentDomain.BaseDirectory + "WORK";

            //設定ファイル存在チェック
            if (!System.IO.File.Exists(ConfigFilePass.Content.ToString()))
            {
                MessageBox.Show("設定ファイルTTSVConfig.txtが実行ファイルのフォルダにありません。TTSVを終了します。");
                this.Close();
            };

            //設定ファイル読み込み
            TTSVConfig TTSVCFG = new TTSVConfig(ConfigFilePass.Content.ToString());

            //設定ファイルの項目設定
            OpenJTalkOption.Content        = TTSVCFG.OpenJTalkOption;
            compID.Content                 = TTSVCFG.compID;
            vVoiceName.Content             = TTSVCFG.vVoiceName;
            Comment.Content                = TTSVCFG.Comment;
            BaseNoteNumValue.Content       = TTSVCFG.BaseNoteNum;
            MiddleNoteNumValue.Content     = TTSVCFG.MiddleNoteNum;
            TopNoteNumValue.Content        = TTSVCFG.TopNoteNum;
            NaturalEndNoteNumValue.Content = TTSVCFG.NaturalEndNoteNum;
            BottomNoteNumValue.Content     = TTSVCFG.BottomNoteNum;
            QuestionNoteNumValue.Content   = TTSVCFG.QuestionNoteNum;
            NoteSplitMode.Content          = TTSVCFG.NoteSplitMode;
            PrefixValue.Content            = TTSVCFG.Prefix;
            OpenJTalkOption.Content        = TTSVCFG.OpenJTalkOption;

            //設定ファイルに従い画面の背景色を変更する
            this.Background = new SolidColorBrush(Color.FromArgb(255,TTSVCFG.ColorR, TTSVCFG.ColorG, TTSVCFG.ColorB)); 
        }

        /// <summary>
        /// テキストボックス
        /// </summary>
        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        /// <summary>
        /// ＶＳＱＸファイル生成ボタン
        /// </summary>
        private void MakeVSQX_Click(object sender, RoutedEventArgs e)
        {
            //ファイル名用の現在時刻タイムスタンプを取得する
        　　string NowString = DateTime.Now.ToString("yyyyMMdd_HHmmss_");

            string EntryString;
            //ファイル名用にテキストボックスの先頭10文字を取得する
            if (EntryText.Text.Length >= 10)
            {
                EntryString = EntryText.Text.Substring(0, 10);
            }
            else
            {
                EntryString = EntryText.Text;
            }

            //設定ファイル読み込み
            TTSVConfig TTSVCFG = new TTSVConfig(ConfigFilePass.Content.ToString());

            //ファイル名フルパスを設定する
            string EntryTextFileName = EntryTextPass.Content+ "\\" + TTSVCFG.Prefix + NowString + EntryString + ".txt";
            string TraceFileName     = WorkFilePass.Content + "\\" + TTSVCFG.Prefix + NowString + "Trace.txt";
            string VSQXFileName      = VSQXPass.Content     + "\\" + TTSVCFG.Prefix + NowString + EntryString + ".vsqx";
            string AnalysisFileName  = WorkFilePass.Content + "\\" + TTSVCFG.Prefix + NowString + "Analysis.xml";
            string LabelFileName     = WorkFilePass.Content + "\\" + TTSVCFG.Prefix + NowString + "Label.xml";

            //入力テキストをファイルに保存する
            StreamWriter swEntry = new StreamWriter(EntryTextFileName, false, Encoding.GetEncoding("Shift_JIS"));
            swEntry.WriteLine(EntryText.Text);
            swEntry.Close();

            //OpenJTalkを起動する
            ProcessStartInfo OJInfo = new ProcessStartInfo();

            OJInfo.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "open_jtalk.exe";
            OJInfo.Arguments = TTSVCFG.OpenJTalkOption + " -ot " + "\"" + TraceFileName + "\"" + " " + "\"" + EntryTextFileName + "\"";    // コマンドパラメータ（引数）
            OJInfo.CreateNoWindow  = true;    // コンソール・ウィンドウを開かない
            OJInfo.UseShellExecute = false;  // シェル機能を使用しない

            Process OJP = Process.Start(OJInfo);
            OJP.WaitForExit();  // プロセスの終了を待つ

            //TI2VSQXを起動する
            ProcessStartInfo TIInfo = new ProcessStartInfo();
            TIInfo.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "TI2VSQX.exe";
            TIInfo.Arguments = "\"" + ConfigFilePass.Content + "\"" + " " +
                               "\"" + TraceFileName          + "\"" + " " +
                               "\"" + AnalysisFileName       + "\"" + " " +
                               "\"" + LabelFileName          + "\"" + " " +
                               "\"" + VSQXFileName           + "\"";
            TIInfo.CreateNoWindow  = true;    // コンソール・ウィンドウを開かない
            TIInfo.UseShellExecute = false;  // シェル機能を使用しない

            Process TIP = Process.Start(TIInfo);
            TIP.WaitForExit();  // プロセスの終了を待つ

            MessageBox.Show("ＶＳＱＸファイル生成しました");

        }

        /// <summary>
        /// 設定ファイル読み込みボタン
        /// </summary>
        private void SelectConfig_Click(object sender, RoutedEventArgs e)
        {
            // ダイアログのインスタンスを生成
            var dialog = new OpenFileDialog();

            // ファイルの種類を設定
            dialog.Filter = "テキストファイル (*.txt)|*.txt";

            // ダイアログを表示する
            if (dialog.ShowDialog() == true)
            {
                // 選択されたファイル名 (ファイルパス) をラベルに表示
                ConfigFilePass.Content = dialog.FileName;
            }

            //設定ファイル読み込み
            TTSVConfig TTSVCFG = new TTSVConfig(ConfigFilePass.Content.ToString());

            //設定ファイルの項目設定
            OpenJTalkOption.Content        = TTSVCFG.OpenJTalkOption;
            compID.Content　　　　　       = TTSVCFG.compID;
            vVoiceName.Content             = TTSVCFG.vVoiceName;
            Comment.Content                = TTSVCFG.Comment;
            BaseNoteNumValue.Content       = TTSVCFG.BaseNoteNum;
            MiddleNoteNumValue.Content     = TTSVCFG.MiddleNoteNum;
            TopNoteNumValue.Content        = TTSVCFG.TopNoteNum;
            NaturalEndNoteNumValue.Content = TTSVCFG.NaturalEndNoteNum;
            BottomNoteNumValue.Content     = TTSVCFG.BottomNoteNum;
            QuestionNoteNumValue.Content   = TTSVCFG.QuestionNoteNum;
            NoteSplitMode.Content          = TTSVCFG.NoteSplitMode;
            PrefixValue.Content            = TTSVCFG.Prefix;
            OpenJTalkOption.Content        = TTSVCFG.OpenJTalkOption;

            //設定ファイルに従い画面の背景色を変更する
            this.Background = new SolidColorBrush(Color.FromArgb(255, TTSVCFG.ColorR, TTSVCFG.ColorG, TTSVCFG.ColorB));
        }

        /// <summary>
        /// 入力テキスト保存先フォルダ設定ボタン
        /// </summary>
        private void SetEntryPass_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();

            dlg.Title                     = "入力テキスト保存先フォルダの選択";
            dlg.IsFolderPicker            = true;
            dlg.InitialDirectory          = EntryTextPass.Content.ToString();
            dlg.AddToMostRecentlyUsedList = false;
            dlg.AllowNonFileSystemItems   = false;
            dlg.DefaultDirectory          = EntryTextPass.Content.ToString();
            dlg.EnsureFileExists          = true;
            dlg.EnsurePathExists          = true;
            dlg.EnsureReadOnly            = false;
            dlg.EnsureValidNames          = true;
            dlg.Multiselect               = false;
            dlg.ShowPlacesList            = true;

            if (dlg.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                EntryTextPass.Content = dlg.FileName;
            }
        }

        /// <summary>
        /// ＶＳＱＸファイル保存先フォルダ設定ボタン
        /// </summary>
        private void SetVSQXPass_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();

            dlg.Title                     = "ＶＳＱＸファイル保存先フォルダの選択";
            dlg.IsFolderPicker            = true;
            dlg.InitialDirectory          = EntryTextPass.Content.ToString();
            dlg.AddToMostRecentlyUsedList = false;
            dlg.AllowNonFileSystemItems   = false;
            dlg.DefaultDirectory          = EntryTextPass.Content.ToString();
            dlg.EnsureFileExists          = true;
            dlg.EnsurePathExists          = true;
            dlg.EnsureReadOnly            = false;
            dlg.EnsureValidNames          = true;
            dlg.Multiselect               = false;
            dlg.ShowPlacesList            = true;

            if (dlg.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                VSQXPass.Content = dlg.FileName;
            }
        }

        /// <summary>
        /// 作業ファイル保存先フォルダ設定ボタン
        /// </summary>
        private void SetWorkFilePass_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();

            dlg.Title                     = "作業ファイル保存先フォルダの選択";
            dlg.IsFolderPicker            = true;
            dlg.InitialDirectory          = EntryTextPass.Content.ToString();
            dlg.AddToMostRecentlyUsedList = false;
            dlg.AllowNonFileSystemItems   = false;
            dlg.DefaultDirectory          = EntryTextPass.Content.ToString();
            dlg.EnsureFileExists          = true;
            dlg.EnsurePathExists          = true;
            dlg.EnsureReadOnly            = false;
            dlg.EnsureValidNames          = true;
            dlg.Multiselect               = false;
            dlg.ShowPlacesList            = true;

            if (dlg.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                WorkFilePass.Content = dlg.FileName;
            }
        }

        /// <summary>
        /// 画面終了
        /// </summary>
        private void exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
