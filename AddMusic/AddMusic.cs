using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

using System.Drawing;

namespace AddMusic
{

    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            if (vegas.Project.Tracks.Count == 0)
            {
                // トラックがないときは何もしない
                MessageBox.Show("Error: No tracks");
                return;
            }

            // パラメータを設定するダイアログを開く
            Tuple<DialogResult, string, string> dialogResult = DoDialog(vegas);
            DialogResult result = dialogResult.Item1;
            if (result != DialogResult.OK)
            {
                return;
            }

            string VideoTrackNo1BaseString = dialogResult.Item2;
            string MusicTrackNo1BaseString = dialogResult.Item3;

            // ダイアログを開き出力するログのファイルパスをユーザーに選択させる
            string saveFilePath = GetFilePath(vegas.Project.FilePath, "AddMusic");
            if (saveFilePath.Length == 0)
            {
                MessageBox.Show("Error: No file path");
                return;
            }

            System.IO.StreamWriter writer = new System.IO.StreamWriter(saveFilePath, false, Encoding.GetEncoding("Shift_JIS"));

            MediaPool mediaPool = vegas.Project.MediaPool;
            AudioFiles audioFiles = new AudioFiles(mediaPool);
            if (audioFiles == null)
            {
                MessageBox.Show("Error: Internal error. Can't create AudioFiles");
                return;
            }

            //            audioFiles.Show(writer);


 
            int VideoTrackNo1Base = 1;
            if (!int.TryParse(VideoTrackNo1BaseString, out VideoTrackNo1Base))
            {
                MessageBox.Show("Cannot parse VideoTrackNoString");
                return;
            }
            if (VideoTrackNo1Base == 0)
            {
                MessageBox.Show("VideoTrackNoString must not be 0.");
                return;
            }
            int VideoTrackNo0Base = VideoTrackNo1Base - 1;

            int MusicTrackNo1Base = 4;
            if (!int.TryParse(MusicTrackNo1BaseString, out MusicTrackNo1Base))
            {
                MessageBox.Show("Cannot parse MusicTrackNoString");
                return;
            }
            if (MusicTrackNo1Base == 0)
            {
                MessageBox.Show("MusicTrackNoString must not be 0.");
                return;
            }

            int MusicTrackNo0Base = MusicTrackNo1Base - 1;

            if ((vegas.Project.Tracks.Count <= VideoTrackNo0Base) || (vegas.Project.Tracks.Count <= MusicTrackNo0Base))
            {
                MessageBox.Show("Error: Requires more tracks.");
                return;
            }

            Track videoTrack = vegas.Project.Tracks[VideoTrackNo0Base];
            Track musicTrack = vegas.Project.Tracks[MusicTrackNo0Base];

            //{

            //    // music Trackの内容を出力する
            //    foreach (TrackEvent trackEvent in musicTrack.Events)
            //    {
            //        string str = trackEvent.Start.ToString() + " " + trackEvent.SnapOffset.ToString() + " " + trackEvent.ActiveTake.Offset.ToString() + " " + trackEvent.ActiveTake.MediaStream.Offset.ToString();
            //        writer.WriteLine(str);
            //    }

            //    writer.Close();
            //    return;
            //}



            // 画像がこの枚数続いたときに音楽を追加する
            const int ConsecutiveImagesThreshold = 2;
            MusicAdder musicAdder = new MusicAdder(audioFiles, musicTrack, ConsecutiveImagesThreshold);
            if (musicAdder == null)
            {
                MessageBox.Show("Error: Internal error. Can't create MusicAdder");
                return;
            }

            long err = 0;
            foreach (TrackEvent trackEvent in videoTrack.Events)
            {

                Take take = trackEvent.ActiveTake;
                if (take == null)
                {
                    continue;
                }

                Media media = take.Media;
                string path = media.FilePath; // メディアのファイルパス

                if (!IsImage(path))
                {
                    // Event は Image ではない
                    err = musicAdder.FoundExceptImage();
                    if (err != 0)
                    {
                        break;
                    }

                    continue;
                }

                // Event が Image ならここに来る
                err = musicAdder.FoundImage(trackEvent.Start, trackEvent.End);
                if (err != 0)
                {
                    break;
                }
            }

            err = musicAdder.Finished();
            if (err != 0)
            {
            }

            writer.Close();

        }

        // ダイアログを開きファイルパスをユーザーに選択させる
        private string GetFilePath(string rootFilePath, string preFix)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = preFix + System.IO.Path.GetFileNameWithoutExtension(rootFilePath) + ".txt";
            sfd.InitialDirectory = System.IO.Path.GetDirectoryName(rootFilePath) + "\\";
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return "";
            }

            return sfd.FileName;
        }

        // 与えられたファイルパスの拡張子が、extentionsに含まれている拡張子であるかを調べる
        private bool isSameExtention(string path, string[] extentions)
        {
            // ファイルパスから拡張子を取得
            string searchExtention = Path.GetExtension(path).ToLower();

            // extentionsに含まれている拡張子かを調べる
            foreach (string extention in extentions)
            {
                if (extention == searchExtention)
                {
                    return true;
                }
            }

            return false;
        }

        // 画像のファイルパスかどうかを返す
        // 画像かどうかは、拡張子で判断する
        private bool IsImage(string path)
        {
            // 画像の拡張子リスト
            string[] Extentions = new string[] { ".jpg", ".jpeg", ".heic" };

            return isSameExtention(path, Extentions);
        }


        private Tuple<DialogResult, string, string> DoDialog(Vegas vegas)
        {
            Form form = new Form();
            form.SuspendLayout();
            form.AutoScaleMode = AutoScaleMode.Font;
            form.AutoScaleDimensions = new SizeF(6F, 13F);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.HelpButton = false;
            form.ShowInTaskbar = false;
            form.Text = "Add Music";
            form.AutoSize = true;
            form.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.AutoSize = true;
            layout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            layout.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            layout.ColumnCount = 3;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            form.Controls.Add(layout);

            {
                Label label;

                label = new Label();
                label.Text = "Video Track No(1-base):";
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Margin = new Padding(8, 8, 8, 4);
                label.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                layout.Controls.Add(label);
                layout.SetColumnSpan(label, 3);
            }

            TextBox VideoTrackNoBox = new TextBox();
            {
                VideoTrackNoBox.Text = "2";

                VideoTrackNoBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                VideoTrackNoBox.Margin = new Padding(16, 8, 8, 4);
                layout.Controls.Add(VideoTrackNoBox);
                layout.SetColumnSpan(VideoTrackNoBox, 2);
            }

            {
                Label label;

                label = new Label();
                label.Text = "Music Track No(1-base):";
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Margin = new Padding(8, 8, 8, 4);
                label.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                layout.Controls.Add(label);
                layout.SetColumnSpan(label, 3);
            }

            TextBox MusicTrackNoBox = new TextBox();
            {
                MusicTrackNoBox.Text = "5";

                MusicTrackNoBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                MusicTrackNoBox.Margin = new Padding(16, 8, 8, 4);
                layout.Controls.Add(MusicTrackNoBox);
                layout.SetColumnSpan(MusicTrackNoBox, 2);
            }



            {
                FlowLayoutPanel buttonPanel = new FlowLayoutPanel();
                buttonPanel.FlowDirection = FlowDirection.RightToLeft;
                buttonPanel.Size = Size.Empty;
                buttonPanel.AutoSize = true;
                buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                buttonPanel.Margin = new Padding(8, 8, 8, 8);
                buttonPanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                layout.Controls.Add(buttonPanel);
                layout.SetColumnSpan(buttonPanel, 3);

                {
                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.FlatStyle = FlatStyle.System;
                    cancelButton.DialogResult = DialogResult.Cancel;
                    buttonPanel.Controls.Add(cancelButton);
                    form.CancelButton = cancelButton;
                }

                {
                    Button okButton = new Button();

                    okButton.Text = "OK";
                    okButton.FlatStyle = FlatStyle.System;
                    okButton.DialogResult = DialogResult.OK;
                    buttonPanel.Controls.Add(okButton);
                    form.AcceptButton = okButton;
                }
            }

            form.ResumeLayout();

            string textVideoTrackNo = "";
            string textMusicTrackNo = "";
            DialogResult result = form.ShowDialog(vegas.MainWindow);
            if (DialogResult.OK == result)
            {
                textVideoTrackNo = VideoTrackNoBox.Text;
                textMusicTrackNo = MusicTrackNoBox.Text;
            }

            return Tuple.Create(result, textVideoTrackNo, textMusicTrackNo);
        }

    }

}

