namespace HomeFront.Components.Hf;

/// <summary>
/// Interface for HfGrid to receive column registration from HfGridColumnsBase.
/// </summary>
internal interface IHfGridOwner
{
    void RegisterColumnsContainer(HfGridColumnsBase container);
    void NotifyColumnsChanged();
}
