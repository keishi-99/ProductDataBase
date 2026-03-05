namespace ProductDatabase {
    public class ListItem<T> {
        public T Id { get; set; } = default!;
        public string Name { get; set; } = string.Empty;

        public override string ToString() => Name;
    }
}
