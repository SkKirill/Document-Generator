namespace CreaterFromVSU.ViewModel.Utilites.Structs
{
    public struct PlayersListStruct
    {
        public string? CodeCompetition;
        public string? NameCommand;
        public string? eMail;
        public string? CodeExhibition;
        public string? CodeContest;
        public string? OlympicsContest;
        public string FioPlayers;
        public DateTime BirthdayPlayers;
        public bool isMen;
        public string SchoolPlayers;
        public string? CityPlayers;
        public string? TeacherPlayers;

        public override string ToString()
        {
            return FioPlayers.PadRight(35, ' ') + " | " + CityPlayers;
        }
    }
}
