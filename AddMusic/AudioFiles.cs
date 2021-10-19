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

        // �w�肳�ꂽ�������̉��y��Ԃ�
        // �߂�l�F�iaudioStream�AaudioStream�̊J�n�����AaudioStream�̒����j�̑g
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
                    // currentStream �̌��݈ʒu����K�v�Ȓ��� length ������Ƃ�
                    audioStreams.Add(Tuple.Create(currentStream, currentTime_, remainLength));
                    currentTime_ += remainLength;
                    usedTime += remainLength;
                }
                else
                {
                    // currentStream �̌��݈ʒu����K�v�Ȓ��� length �����Ȃ�

                    // �܂��AcurrentStream ���Ō�܂Ŏg��
                    Timecode thisLength = currentStream.Length - currentTime_;
                    audioStreams.Add(Tuple.Create(currentStream, currentTime_, thisLength));
                    usedTime += thisLength;
                    remainLength -= thisLength;

                    // ���� stream ���擾����
                    if (currentStreamEnumerator_.MoveNext())
                    {
                        currentTime_ = new Timecode(0);

                    }
                    else
                    {
                        // �����Ȃ��I
                        // MediaStream ������Ȃ��̂Ń��[�U�[�ɁAMediaStream ��ǉ�����悤�Ƀ��b�Z�[�W���o��
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
                string path = media.FilePath; // ���f�B�A�̃t�@�C���p�X

                MediaStream videoStream = media.GetVideoStreamByIndex(0);
                if (videoStream != null)
                {
                    // Video ����ꂽ��Ⴄ
                    continue;
                }

                MediaStream audioStream = media.GetAudioStreamByIndex(0);
                if (audioStream == null)
                {
                    // Audio �����Ȃ���ΈႤ
                    continue;
                }

                audioStreams.Add(audioStream);
            }

            return 0;
        }
    }

}
