using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace AddMusic
{
    public class MusicAdder
    {
        private AudioFiles audioFiles_;
        private Track musicTrack_;  // ���y��ǉ�����g���b�N
        private int consecutiveImagesThreshold_ = 0;    // 
        private int consecutiveImagesNo_ = 0;
        private Timecode startTime_;
        private Timecode endTime_;
        private bool musicStarting_ = false;

        private TrackEvent audioEvent_;
        private Timecode addedStreamLength_;

        // 2 �̉摜�𑱂��Ĕz�u�����Ƃ��̃I�[�o�[���b�v���ԁ@�i�u���[�U�[�ݒ�v->�u�ҏW�v->�u�I�[�o�[���b�v�̃J�b�g�ϊ��̕b���j
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
                startTime_ = startTime + ImageOverlapLength_;   // �I�[�o�[���b�v���Ă���Ԃ͉��y�͗����Ȃ�
                endTime_ = endTime - ImageOverlapLength_;   // �I�[�o�[���b�v���Ă���Ԃ͉��y�͗����Ȃ�
            }

            consecutiveImagesNo_++;

            if (IsStartingMusic())
            {
                // ���łɉ��y���n�܂��Ă�����A���y�𑱂���
                err = ContinueMusic(endTime);
            }
            else
            {
                if (CanStartMusic())
                {
                    // �摜���K�萔�ȏ�A�����Ă���̂ŁA���y��ǉ�����
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

            // �g���Ȃ������C�x���g�͏���
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
            endTime_ = endTime - ImageOverlapLength_;   // �I�[�o�[���b�v���Ă���Ԃ͉��y�͗����Ȃ�
            musicStarting_ = true;

            return 0;
        }

        private long ContinueMusic(Timecode endTime)
        {
            // ���y�̏I���������X�V����
            endTime_ = endTime - ImageOverlapLength_;   // �I�[�o�[���b�v���Ă���Ԃ͉��y�͗����Ȃ�

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

            // �ǉ����鉹�y�̒������Y�o����
            Timecode length = endTime_ - startTime_;

            // �Y�o�����������̉��y���擾
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

            // �擾�������y���A�C�x���g�̊J�n�����ɒ���t����
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
                    // �ŏ��� 1 �߂́A���y�t�@�C���̒������̃C�x���g��ǉ�����
                    audioEvent_ = new AudioEvent(startTime, audioLength);
                    musicTrack_.Events.Add(audioEvent_);

                    Take take = new Take(mediaStream);
                    audioEvent_.Takes.Add(take);

                    addedStreamLength_.FrameCount = 0;

                    isFirstClip = true;
                }
                else
                {
                    // 2 �߈ȍ~�́A�C�x���g�̊J�n�ʒu�����炷
                    audioEvent_.AdjustStartLength(startTime, audioLength - addedStreamLength_, true);
                    audioEvent_.ActiveTake.Offset = addedStreamLength_;
                }

                startTime += streamLength;
                addedStreamLength_ += streamLength;
                if (audioLength <= addedStreamLength_)
                {
                    Timecode fadeLength = Timecode.FromString("00:00:01;00");   // Audio �̃t�F�[�h����
                    audioEvent_.FadeIn.Length = fadeLength;
                    // ���y�t�@�C���̍Ō�ɂ� FadeOut �͓���Ȃ�

                    audioEvent_ = null;
                }
                else
                {
                    // mediaStream ��ݒ肵�Ă����Ԃ� Split ����΁ASplit �����O�オ TrackEvent �ɒǉ�����Ă���
                    remainAudioEvent = audioEvent_.Split(streamLength);

                    Timecode fadeLength = Timecode.FromString("00:00:01;00");   // Audio �̃t�F�[�h����
                    if (!isFirstClip)
                    {
                        // ���y�t�@�C���̐擪�ɂ� FadeIn �͓���Ȃ�
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
