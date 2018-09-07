//This code belongs to Adam Korytowski, Kraków, Poland


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Windows.Input;
using WMPLib;
using System.Windows.Controls;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer synthesiser = new SpeechSynthesizer();
        WMPLib.IWMPMedia media = null;

        SpeechRecognitionEngine engine = new SpeechRecognitionEngine();

        Choices choices = new Choices();

        string filePath = @"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\bin\Debug\songNames.txt";
        string nircmdPath = @"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\nircmd\nircmd.exe";

        //inicjalizacja okna graficznego
        public Form1()
        {
            InitializeComponent();
        }
        
        //konstruktor okna graficznego, wywolywany przy uruchamianiu programu
        private void Form1_Load(object sender, EventArgs e)
        {
            AllocConsole();
            killButton.Text = "Close";

            Image image = Image.FromFile(@"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\stopIcon.jpg");
            playPicture.Image = image;

            TopMost = true;

            wmplayer.settings.volume = 50;
            wmplayer.settings.mute = false;

            loadSongs();
            makeUpListbox();


            this.BackColor = Color.FromArgb(171, 0, 42);
            this.Text = "Katherine";
            

            killButton.Visible = false;
            katherineLabel.Text = "Say: Katherine +...";
           

            addCommands_2listbox();
            volProgressBar.Value = Int32.Parse(wmplayer.settings.volume.ToString());
            volLabel.Text = "Vol: ";

            playlistLabel.Text = "Your Playlist: ";;
            playlistLabel.Font = new Font(playlistLabel.Font.FontFamily, 12, FontStyle.Bold);
            playlistLabel.ForeColor = Color.FromArgb(0, 199, 0);
        }

        //dodanie logu programu w formie kolsoli 
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        //dodanie komend do okna, aby uzytkownik widzial jakie komendy glosowe sa dostepne
        private void addCommands_2listbox()
        {
            setUpKatherineLabel();
            commandsBox.Items.Add("play + <song from Playlist>");
            commandsBox.Items.Add("what is this song?");
            commandsBox.Items.Add("volume up");
            commandsBox.Items.Add("volume down");
            commandsBox.Items.Add("volume mute");
            commandsBox.Items.Add("volume unmute");
            commandsBox.Items.Add("play next");
            commandsBox.Items.Add("play previous");
            commandsBox.Items.Add("play");
            commandsBox.Items.Add("pause");
            commandsBox.Items.Add("stop music");
            commandsBox.Items.Add("loop this playlist");
            commandsBox.Items.Add("shuffle playlist");
            commandsBox.Items.Add("system volume up");
            commandsBox.Items.Add("system volume down");
            commandsBox.Items.Add("mute system");
            commandsBox.Items.Add("unmute system");
            commandsBox.Items.Add("remove this song");
        }

        //stworzenie labela informujacego o skladni
        private void setUpKatherineLabel()
        {
            katherineLabel.Text = "Say 'Katherine +...";
            katherineLabel.Font = new Font(katherineLabel.Font.FontFamily, 12, FontStyle.Bold);
            katherineLabel.ForeColor = Color.FromArgb(0, 199, 0);
        }
        
        //stworzenie playlisty po kliknieciu przycisku
        private void button1_Click_1(object sender, EventArgs e)
        {

            if (!string.IsNullOrWhiteSpace(pasteLinkBox.Text))
            {
                choices.Add(new string[] {"Katherine mute","Katherine unmute", 
                "Katherine play No Role Modelz", "Katherine play 2pac changes",
                "Katherine go to sleep","Katherine volume down","Katherine volume up",
                "Katherine volume down 20","Katherine volume up 20","Katherine stop music",
                "Katherine next song","Katherine play previous song","Katherine pause","Katherine play"});

                completeGrammar();

                listBox1.Items.Clear();
                string link = pasteLinkBox.Text.ToString();
                addSongNamesTotheDictionary(choices, link);

                GrammarBuilder grammarBuilder = new GrammarBuilder();
                grammarBuilder.Append(choices);
                Grammar grammar = new Grammar(grammarBuilder);

                engine.LoadGrammarAsync(grammar);
                engine.SetInputToDefaultAudioDevice();
                engine.SpeechRecognized += engine_SpeechRecognized;

                engine.RecognizeAsync(RecognizeMode.Multiple);

                button1.Enabled = false;
            }
        }


        private void recognizer_AudioLevelUpdated(object sender, AudioLevelUpdatedEventArgs e)
        {
            Console.WriteLine("The audio level is now: {0}.", e.AudioLevel);
        }


        private void loadSongs()
        {
            wmplayer.Ctlcontrols.play();
            wmplayer.Ctlcontrols.stop();        //po to zeby piosenki sie wczytaly z dysku
        }

        //usuniecie ".mp3" z nazw utworu, aby nie byly widoczne w playliscie
        private void makeUpListbox()
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
                listBox1.Items[i] = listBox1.Items[i].ToString().Substring(0, listBox1.Items[i].ToString().Length - 4);
        }

        //uzupelnienie gramatyki o niezbedne komendy
        private void completeGrammar()
        {
            choices.Add("Katherine what is this song?");            
            choices.Add("Katherine loop this playlist");
            choices.Add("Katherine shuffle playlist");
            choices.Add("Katherine remove this song");
            choices.Add("Katherine play next");
            choices.Add("Katherine play previous");
            choices.Add("Katherine system volume down");
            choices.Add("Katherine system volume up");

        }

        //funkcja do zmiany ikonki informujace o aktualnym stanie playera
        private void changeState(string state)
        {
            switch (state)
            {
                case "playing":
                    playPicture.Image = Image.FromFile(@"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\pauseIcon.jpg");
                    break;
                case "stopped":
                    playPicture.Image = Image.FromFile(@"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\stopIcon.jpg");
                    break;
                case "paused":
                    playPicture.Image = Image.FromFile(@"C:\Users\Adam\Documents\VisualStudio2013\Projects\WindowsFormsApplication2\WindowsFormsApplication2\playIcon.jpg");
                    break;
            }
        }

        private void killButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        } 

        //dodanie nazw piosenek do gramatyki systemu rozpoznawania, zczytujac nazwy piosenek z pliku, ktory najpierw zostaje stworzony
        private void addSongNamesTotheDictionary(Choices choices_, string LINK)
        {
            Process getSongNamesProcess = new Process();
            Process clearFile = new Process();
            try
            {
                clearFile.StartInfo.FileName = "cmd.exe";
                clearFile.StartInfo.Arguments = "/c TYPE nul > " + filePath;
                clearFile.Start();

                getSongNamesProcess.StartInfo.FileName = "cmd.exe";
                getSongNamesProcess.StartInfo.Arguments = "/c dir /b " + LINK + " >> songNames.txt";
                getSongNamesProcess.Start();
            }
            catch
            {
                MessageBox.Show("The file couldn't have been created");
            }

            WMPLib.IWMPControls3 controls = (WMPLib.IWMPControls3)wmplayer.Ctlcontrols;
            controls.play();

            WMPLib.IWMPPlaylist playlist = wmplayer.playlistCollection.newPlaylist("myplaylist");
           
            try
            {
                using (StreamReader sr = File.OpenText(filePath))
                {
                    string line = null;
                    while ((line = sr.ReadLine()) != null)
                    {
                        media = wmplayer.newMedia(LINK + @"\" + line);
                        line = line.Substring(0, line.Length - 4);
                        listBox1.Items.Add(line);
                        playlist.appendItem(media);
                        choices_.Add("Katherine play " + line); 
                    }
                    wmplayer.currentPlaylist = playlist;
                    wmplayer.Ctlcontrols.stop();
                }
            }
            catch
            {
                MessageBox.Show("STREAMREADER FAILED ;___;");
            }
        }

        //przeskalowanie funkcji zmieny glosnosci systemowej, gdyz przyjmuje jako argument liczbe 2^16
        int scaledVolume(int volumePercent)
        {
            int scaledValue = 0;
            int upBorder = 65535;
            scaledValue = upBorder / (100 / volumePercent);
            return scaledValue;
        }

        void addInfo2Log(string info)
        {
            Console.WriteLine(info);
        }

        //funkcja odpowiedzialna za zsyntesowanie tytulu piosenki i wykonawcy
        private void answerSongName()
        {
            string currentSongName = wmplayer.currentMedia.getItemInfo("Title").ToString();    //odczytanie ostatniego odtwarzanego utworu

            string artist = wmplayer.currentMedia.getItemInfo("Artist");

            if (wmplayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {  
                synthesiser.SpeakAsync(currentSongName + " by " + artist);
                addInfo2Log("song answered: " + currentSongName + " by " + artist);
            }
            else
            {
                synthesiser.SpeakAsync("Music not currently playing. Previous playing song was:" + 
                    currentSongName + " by " + artist);
            }
        }

        //funkcja odpowiedzialna za usuniecie piosenki z playlisty
        private void removeSong()
        {
            string songName = wmplayer.currentMedia.getItemInfo("Title");
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                if (listBox1.Items[i].ToString().Contains(songName))
                {
                    listBox1.Items.Remove(songName);
                }
            }
            if (wmplayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                wmplayer.currentPlaylist.removeItem(wmplayer.currentMedia);     //to dziala, usuwa utwor ktory chce usunac 
                wmplayer.Ctlcontrols.play();
                Console.WriteLine(songName + " was removed from playlist.");
            }
            else synthesiser.SpeakAsync("Music not currently playing"); 
        }

        //najwazniejsza funkcja, odpowiedzialna za przypisanie komend do wykonywanych akcji
        private void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            if (e.Result.Confidence>=0.9)
            {
                Console.Write("Reconized ~ " + e.Result.Text + " ~ with confidence " + e.Result.Confidence + '\n' );

                switch (e.Result.Text)
                {
                    case "Katherine mute system":
                        Process mute_process = new Process();
                        mute_process.StartInfo.FileName = nircmdPath;
                        mute_process.StartInfo.Arguments = "mutesysvolume 1";
                        mute_process.Start();
                        Console.WriteLine("system muted");
                        break;
                    case "Katherine unmute system":
                        Process unmute_process = new Process();
                        unmute_process.StartInfo.FileName = nircmdPath;
                        unmute_process.StartInfo.Arguments = "mutesysvolume 0";
                        unmute_process.Start();
                        Console.WriteLine("system unmuted");
                        break;
                    case "Katherine system volume down":
                        Process vDown = new Process();
                        vDown.StartInfo.FileName = nircmdPath;
                        vDown.StartInfo.Arguments = "changesysvolume " + scaledVolume(-30).ToString();
                        vDown.Start();
                        break;
                    case "Katherine system volume up":
                        Process vUp = new Process();
                        vUp.StartInfo.FileName = nircmdPath;
                        vUp.StartInfo.Arguments = "changesysvolume " + scaledVolume(30).ToString();
                        vUp.Start();
                        break;
                    case "Katherine system volume down 20":
                        Process vDown20 = new Process();
                        vDown20.StartInfo.FileName = nircmdPath;
                        vDown20.StartInfo.Arguments = "changesysvolume " + scaledVolume(-20).ToString();
                        vDown20.Start();
                        break;
                    case "Katherine system volume up 20":
                        Process vUp20 = new Process();
                        vUp20.StartInfo.FileName = nircmdPath;
                        vUp20.StartInfo.Arguments = "changesysvolume " + scaledVolume(20).ToString();
                        vUp20.Start();
                        break;
                    case "Katherine stop music":
                        changeState("stopped");
                        wmplayer.Ctlcontrols.stop();
                        addInfo2Log("Music stopped");
                        break;
                    case "Katherine play next":
                        if (wmplayer.playState != WMPLib.WMPPlayState.wmppsPlaying)
                        {
                            changeState("playing");
                        }
                        wmplayer.Ctlcontrols.next();
                        wmplayer.Ctlcontrols.play();
                        addInfo2Log("Next song playing");
                        Console.WriteLine(wmplayer.currentMedia.getItemInfo("Title") + " playing");
                        break;
                    case "Katherine play previous":
                        if (wmplayer.playState != WMPLib.WMPPlayState.wmppsPlaying)
                        {
                            changeState("playing");
                        }
                        wmplayer.Ctlcontrols.previous();
                        wmplayer.Ctlcontrols.play();
                        addInfo2Log("Previous song playing");
                        Console.WriteLine(wmplayer.currentMedia.getItemInfo("Title") + " playing");
                        break;
                    case "Katherine pause":
                        changeState("paused");
                        wmplayer.Ctlcontrols.pause();
                        addInfo2Log("Music paused");
                        break;
                    case "Katherine play":
                        changeState("playing");
                        wmplayer.Ctlcontrols.play();
                        if (wmplayer.playState != WMPLib.WMPPlayState.wmppsPlaying)
                            Console.WriteLine(wmplayer.currentMedia.getItemInfo("Title") + " playing");
                        break;
                    case "Katherine what is this song?":
                        answerSongName();
                        break;
                    case "Katherine loop this playlist":
                        wmplayer.settings.setMode("loop", true);
                        Console.WriteLine("playlist looped");
                        break;
                    case "Katherine shuffle playlist":
                        wmplayer.settings.setMode("shuffle", true);
                        Console.WriteLine("random songs order made");
                        break;
                    case "Katherine mute":
                        wmplayer.settings.mute = true;
                        Console.WriteLine("player muted");
                        break;
                    case "Katherine unmute":
                        wmplayer.settings.mute = false;
                        Console.WriteLine("player unmuted");
                        break;
                    case "Katherine remove this song":
                        removeSong();
                        break;
                    case "Katherine volume down":
                        wmplayer.settings.volume -= 30;
                        volProgressBar.Value = Int32.Parse(wmplayer.settings.volume.ToString());
                        addInfo2Log("Volume decreased");
                        break;
                    case "Katherine volume up":
                        wmplayer.settings.volume += 30;
                        volProgressBar.Value = Int32.Parse(wmplayer.settings.volume.ToString());
                        addInfo2Log("Volume increased");
                        break;
                }
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    if (e.Result.Text == "Katherine play " + listBox1.Items[i].ToString())
                    {
                        try
                        {
                            changeState("playing");
                            Console.WriteLine(listBox1.Items[i].ToString());
                            media = wmplayer.currentPlaylist.get_Item(i);
                            wmplayer.Ctlcontrols.playItem(media);
                        }
                        catch
                        {
                            MessageBox.Show("song play error");
                        }
                    }
                }
            }
        }  
    }
}
