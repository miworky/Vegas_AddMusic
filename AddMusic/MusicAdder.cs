using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace AddMusic
{
    public class MusicAdder
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


        public MusicAdder(AudioFiles audioFiles, Track musicTrack, int consecutiveImagesThreshold)
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
            Timecode length = endTime_ - startTime_;

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
}
