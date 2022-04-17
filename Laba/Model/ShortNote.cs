namespace Laba.Model
{
    public class ShortNote
    {
        public string Id { get; }
        public string Name { get; }

        public ShortNote(string id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
