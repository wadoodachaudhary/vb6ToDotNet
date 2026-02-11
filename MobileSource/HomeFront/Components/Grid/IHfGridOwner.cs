namespace HomeFront.Components.Grid;

/// <summary>
/// Interface for HfGrid to receive column registration from HfGridColumnsBase.
/// </summary>
internal interface IHfGridOwner
{
    void RegisterColumnsContainer(HfGridColumnsBase container);
}
