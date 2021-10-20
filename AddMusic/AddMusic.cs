using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

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


            // トラックに張り付けてあるメディアをみる
            // 画像が 2 つ以上連続していれば、その区間に音楽をつける
            // 想定しているトラック構成：
            // 0: テロップ
            // 1: ビデオ
            // 2: ビデオの音声（ゲイン補正あり）
            // 3: ビデオの音声（ゲイン補正なし）
            // 4: 音楽
            const int VideoTrackNo = 1;
            const int MusicTrackNo = 4;

            if (vegas.Project.Tracks.Count <= MusicTrackNo)
            {
                MessageBox.Show("Error: Requires 5 tracks or more.");
                return;
            }

            Track videoTrack = vegas.Project.Tracks[VideoTrackNo];
            Track musicTrack = vegas.Project.Tracks[MusicTrackNo];

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
        // HEIFの exif は読み方がわからないので、非対応
        private bool IsImage(string path)
        {
            // 画像の拡張子リスト
            string[] Extentions = new string[] { ".jpg", ".jpeg" };

            return isSameExtention(path, Extentions);
        }

    }

}

