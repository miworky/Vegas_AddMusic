using System;
using System.Collections.Generic;
using ScriptPortal.Vegas;

namespace AddMusic
{

    public class AudioFiles
    {
        private List<MediaStream> audioStreams_;
        private List<MediaStream>.Enumerator currentStreamEnumerator_;
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

}
