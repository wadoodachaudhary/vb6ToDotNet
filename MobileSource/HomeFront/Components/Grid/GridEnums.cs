namespace HomeFront.Components.Grid;

public enum ColumnType
{
    Text,
    Number,
    Date,
    Boolean,
    CheckBox
}

public enum TextAlign
{
    Left,
    Center,
    Right
}

public enum SortDirection
{
    Ascending,
    Descending
}

public enum FilterType
{
    FilterBar,
    Menu,
    CheckBox
}

public enum EditMode
{
    Inline,
    Dialog
}

public enum SelectionType
{
    Single,
    Multiple
}

public enum SelectionMode
{
    Row,
    Cell
}

public enum ClipMode
{
    Clip,
    Ellipsis,
    EllipsisWithTooltip
}

public enum GridLines
{
    Default,
    Both,
    Horizontal,
    Vertical,
    None
}

public enum NewRowPosition
{
    Top,
    Bottom
}

public enum AggregateType
{
    Sum,
    Average,
    Count,
    Min,
    Max
}
