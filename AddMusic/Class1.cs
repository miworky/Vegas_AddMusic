using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace AddMusic
{

    public class AudioFiles
    {
        private List<MediaStream> audioStreams_;
        private List<MediaStream>.Enumerator currentStreamEnumerator_ ;
        private Timecode currentTime_;

        public AudioFiles(MediaPool mediaPool)
        {
            audioStreams_ = new List<MediaStream>();

            long err = ListUpAudioFiles(mediaPool, audioStreams_);
            if (err != 0)
            {
                return;
            }
        }

        // 指定された長さ分の音楽を返す
        // 戻り値：（audioStream、audioStreamの開始時刻、audioStreamの長さ）の組
        public List<Tuple<MediaStream, Timecode, Timecode>> GetAudioStreams(Timecode requiredLength)
        {
            Timecode remainLength = requiredLength;

            List<Tuple<MediaStream, Timecode, Timecode>> audioStreams = new List<Tuple<MediaStream, Timecode, Timecode>>();

            if (audioStreams_.Count == 0)
            {
                return audioStreams;
            }

            if (currentStreamEnumerator_.Current == null)
            {
                currentStreamEnumerator_ = audioStreams_.GetEnumerator();
                currentStreamEnumerator_.MoveNext();
                currentTime_ = new Timecode(0);
            }

            Timecode usedTime = new Timecode(0);
            while (usedTime < requiredLength)
            {
                MediaStream currentStream = currentStreamEnumerator_.Current;
                if (currentTime_ + remainLength < currentStream.Length)
                {
                    // currentStream の現在位置から必要な長さ length が取れるとき
                    audioStreams.Add(Tuple.Create(currentStream, currentTime_, remainLength));
                    currentTime_ += remainLength;
                    usedTime += remainLength;
                }
                else
                {
                    // currentStream の現在位置から必要な長さ length が取れない

                    // まず、currentStream を最後まで使う
                    Timecode thisLength = currentStream.Length - currentTime_;
                    audioStreams.Add(Tuple.Create(currentStream, currentTime_, thisLength));
                    usedTime += thisLength;
                    remainLength -= thisLength;

                    // 次の stream を取得する
                    if (currentStreamEnumerator_.MoveNext())
                    {
                        currentTime_ = new Timecode(0);

                    }
                    else
                    {
                        // 次がない！
                        // MediaStream が足りないのでユーザーに、MediaStream を追加するようにメッセージを出す
                        return null;
                    }
                }


            }

            return audioStreams;
        }

        public long Show(System.IO.StreamWriter writer)
        {
            foreach (var audioStreams in audioStreams_)
            {
                string text = audioStreams.StreamInfoText;
                Timecode length = audioStreams.Length;

                writer.WriteLine(text + " " + length);
            }

            return 0;
        }

        private long ListUpAudioFiles(MediaPool mediaPool, List<MediaStream> audioStreams)
        {
            if (mediaPool == null)
            {
                return -1;
            }

            foreach (Media media in mediaPool)
            {
                string path = media.FilePath; // メディアのファイルパス

                MediaStream videoStream = media.GetVideoStreamByIndex(0);
                if (videoStream != null)
                {
                    // Video が取れたら違う
                    continue;
                }

                MediaStream audioStream = media.GetAudioStreamByIndex(0);
                if (audioStream == null)
                {
                    // Audio が取れなければ違う
                    continue;
                }

                audioStreams.Add(audioStream);
            }

            return 0;
        }
    }

    public class AddMusic
    {
        private AudioFiles audioFiles_;
        private Track musicTrack_;  // 音楽を追加するトラック
        private int consecutiveImagesThreshold_ = 0;    // 
        private int consecutiveImagesNo_ = 0;
        private Timecode startTime_;
        private Timecode endTime_;
        private bool musicStarting_ = false;

        private TrackEvent audioEvent_;
        private Timecode addedStreamLength_;

        // 2 つの画像を続けて配置したときのオーバーラップ時間　（「ユーザー設定」->「編集」->「オーバーラップのカット変換の秒数）
        private Timecode ImageOverlapLength_ = Timecode.FromString("00:00:01;00"); 


        public AddMusic(AudioFiles audioFiles, Track musicTrack, int consecutiveImagesThreshold)
        {
            audioFiles_ = audioFiles;
            musicTrack_ = musicTrack;
            consecutiveImagesThreshold_ = consecutiveImagesThreshold;
        }

        public long FoundImage(Timecode startTime, Timecode endTime)
        {
            long err = 0;

            if (consecutiveImagesNo_ == 0)
            {
                startTime_ = startTime + ImageOverlapLength_;   // オーバーラップしている間は音楽は流さない
                endTime_ = endTime - ImageOverlapLength_;   // オーバーラップしている間は音楽は流さない
            }

            consecutiveImagesNo_++;

            if (IsStartingMusic())
            {
                // すでに音楽が始まっていたら、音楽を続ける
                err = ContinueMusic(endTime);
            }
            else
            {
                if (CanStartMusic())
                {
                    // 画像が規定数以上連続しているので、音楽を追加する
                    err = StartMusic(startTime, endTime);
                }
            }

            return err;
        }

        public long FoundExceptImage()
        {
            if (IsStartingMusic())
            {
                long err = StopMusic();
                if (err != 0)
                {
                    return err;
                }
            }

            consecutiveImagesNo_ = 0;

            return 0;
        }

        public long Finished()
        {
            long err = FoundExceptImage();
            if (err != 0)
            {
                return err;
            }

            // 使われなかったイベントは消す
            err = RemoveLastAudioEvent();
            if (err != 0)
            {
                return err;
            }

            return err;
        }

        private bool IsStartingMusic()
        {
            return musicStarting_;
        }

        private bool CanStartMusic()
        {
            if (consecutiveImagesThreshold_ == consecutiveImagesNo_)
            {
                return true;
            }

            return false;
        }

        private long RemoveLastAudioEvent()
        {
            if (audioEvent_ == null)
            {
                return 0;
            }

            if (musicTrack_ == null)
            {
                return -1;
            }

            musicTrack_.Events.Remove(audioEvent_);
            audioEvent_ = null;

            return 0;
        }

        private long StartMusic(Timecode startTime, Timecode endTime)
        {
            endTime_ = endTime - ImageOverlapLength_;   // オーバーラップしている間は音楽は流さない
            musicStarting_ = true;

            return 0;
        }

        private long ContinueMusic(Timecode endTime)
        {
            // 音楽の終了時刻を更新する
            endTime_ = endTime - ImageOverlapLength_;   // オーバーラップしている間は音楽は流さない

            return 0;
        }


        private long StopMusic()
        {
            if (startTime_ == null)
            {
                return -1;
            }

            if (endTime_ == null)
            {
                return -1;
            }

            Timecode startTime = startTime_;

            // 追加する音楽の長さを産出する
            Timecode length = endTime_  - startTime_;

            // 産出した長さ分の音楽を取得
            List<Tuple<MediaStream, Timecode, Timecode>> audioStreams = audioFiles_.GetAudioStreams(length);
            if (audioStreams == null)
            {
                MessageBox.Show("error: Shortage of audio media stream. Please add aditional audio media.");
                return -1;
            }

            if (addedStreamLength_ == null)
            {
                addedStreamLength_ = new Timecode(0);
            }

            // 取得した音楽を、イベントの開始時刻に張り付ける
            foreach (var audioStream in audioStreams)
            {
                MediaStream mediaStream = audioStream.Item1;
                Timecode startTimeOfMediaStream = audioStream.Item2;
                Timecode streamLength = audioStream.Item3;

                TrackEvent remainAudioEvent;

                Timecode audioLength = mediaStream.Length;
                bool isFirstClip = false;
                if (audioEvent_ == null)
                {
                    // 最初の 1 つめは、音楽ファイルの長さ分のイベントを追加する
                    audioEvent_ = new AudioEvent(startTime, audioLength);
                    musicTrack_.Events.Add(audioEvent_);

                    Take take = new Take(mediaStream);
                    audioEvent_.Takes.Add(take);

                    addedStreamLength_.FrameCount = 0;

                    isFirstClip = true;
                }
                else
                {
                    // 2 つめ以降は、イベントの開始位置をずらす
                    audioEvent_.AdjustStartLength(startTime, audioLength - addedStreamLength_, true);
                    audioEvent_.ActiveTake.Offset = addedStreamLength_; 
                }

                startTime += streamLength;
                addedStreamLength_ += streamLength;
                if (audioLength <= addedStreamLength_)
                {
                    Timecode fadeLength = Timecode.FromString("00:00:01;00");   // Audio のフェード時間
                    audioEvent_.FadeIn.Length = fadeLength;
                    // 音楽ファイルの最後には FadeOut は入れない

                    audioEvent_ = null;
                }
                else
                {
                    // mediaStream を設定してある状態で Split すれば、Split した前後が TrackEvent に追加されている
                    remainAudioEvent = audioEvent_.Split(streamLength);

                    Timecode fadeLength = Timecode.FromString("00:00:01;00");   // Audio のフェード時間
                    if (!isFirstClip)
                    {
                        // 音楽ファイルの先頭には FadeIn は入れない
                        audioEvent_.FadeIn.Length = fadeLength;
                    }

                    audioEvent_.FadeOut.Length = fadeLength;

                    audioEvent_ = remainAudioEvent;
                }
            }

            musicStarting_ = false;

            return 0;
        }
    }

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
                MessageBox.Show("Error: Requires 4 tracks or more.");
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
            AddMusic addMusic = new AddMusic(audioFiles, musicTrack, ConsecutiveImagesThreshold);
            if (addMusic == null)
            {
                MessageBox.Show("Error: Internal error. Can't create AddMusic");
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
                    err = addMusic.FoundExceptImage();
                    if (err != 0)
                    {
                        break;
                    }

                    continue;
                }

                // Event が Image ならここに来る
                err = addMusic.FoundImage(trackEvent.Start, trackEvent.End);
                if (err != 0)
                {
                    break;
                }
            }

            err = addMusic.Finished();
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

