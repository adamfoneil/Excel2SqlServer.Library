using System.Threading.Tasks;

namespace Excel2SqlServer.Download.Interfaces
{
    /// <summary>
    /// generalizes a way to store segments of a large result set that can be assembled later
    /// </summary>    
    public interface ISegmentStore<TData, TOperation>
    {
        Task<TOperation> GetOperationIdAsync();
        Task AppendAsync(TOperation operationId, TData data);
        Task<TData> AssembleAsync(TOperation operationId, int skip = 0, int take = 0);
        Task CleanupAsync(TOperation operationId);
        Task<int> GetSegmentCountAsync(TOperation operationId);
    }
}
