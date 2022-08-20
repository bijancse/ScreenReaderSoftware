using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Speech;
using System.Speech.Synthesis;

namespace SpeechBuilder
{
    public class SpeechControl
    {
        SpeechSynthesizer speaker;
       
        public SpeechControl()
        {
            speaker = new SpeechSynthesizer();
            speaker.Rate = 0;
            speaker.Volume = 70;
        }           
       
        public void speak(String word)
        {
            if (word == null || word.Equals("")) return;
            speaker.SpeakAsync(word);
        }
     
        public void stop()
        {

            if (speaker.State == SynthesizerState.Speaking)
            {
                speaker.SpeakAsyncCancelAll();
            }
        }

    }
}
