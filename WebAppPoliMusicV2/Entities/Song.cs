using System;

namespace WebAppPoliMusicV2.Entities
{
    /// <summary>
    /// Entity that maps to the [TBL_SONG] table in the database.
    /// </summary>
    public class Song
    {
        public Song()
        {
        }

        public Song(int id, string name, string path, int plays)
        {
            ID = id;
            Name = name;
            Path = path;
            Plays = plays;
        }

        /// <summary>Maps to [ID_SONG]</summary>
        public int ID { get; set; }

        /// <summary>Maps to [SONG_NAME]</summary>
        public string Name { get; set; }

        /// <summary>Maps to [SONG_PATH]</summary>
        public string Path { get; set; }

        /// <summary>Maps to [PLAYS]</summary>
        public int Plays { get; set; }
    }
}
