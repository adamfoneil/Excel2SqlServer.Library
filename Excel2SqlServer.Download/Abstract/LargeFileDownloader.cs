using ClosedXML.Excel;
using Excel2SqlServer.Download.Interfaces;
using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;

namespace Excel2SqlServer.Download.Abstract
{
    public abstract class LargeFileDownloader<TOperation>
    {
        private readonly ISegmentStore<DataTable, TOperation> _segmentStore;        

        public LargeFileDownloader(ISegmentStore<DataTable, TOperation> segmentStore)
        {
            _segmentStore = segmentStore;
        }

        public abstract int PageSize { get; }

        public async Task<TOperation> BeginAsync()
        {
            var operationId = await _segmentStore.GetOperationIdAsync();
            await InitializeAsync(operationId);
            return operationId;
        }

        /// <summary>
        /// call this from your public endpoint to get the next segment of data and store it for assembly during the CompleteAsync method
        /// </summary>
        public async Task<(bool anyData, int rowCount)> ContinueAsync(TOperation operationId, int pageNumber)
        {
            var result = await QuerySegmentAsync(pageNumber, PageSize);

            if (result.rowCount > 0)
            {
                await _segmentStore.AppendAsync(operationId, result.data);
                return (true, result.rowCount);
            }

            // there are no more segments, time to call CompleteAsync
            return (false, 0);
        }

        public async Task<Stream> CompleteAsync(TOperation operationId, bool forceZip = false, Action<IXLWorksheet> formatWorksheet = null)
        {
            var segmentCount = await _segmentStore.GetSegmentCountAsync(operationId);

            if (segmentCount >= MinZipFileSegments || forceZip)
            {
                var result = new MemoryStream();
                await BuildZipFileAsync(result, segmentCount, operationId, formatWorksheet);
                return result;
            }
            else
            {
                var data = await _segmentStore.AssembleAsync(operationId);
                return BuildWorkbookStream(data, formatWorksheet);
            }
        }

        /// <summary>
        /// minimum number of segments required for the CompleteAsync method to return a zip file
        /// </summary>
        protected abstract int MinZipFileSegments { get; }

        protected abstract int SegmentsPerZipEntry { get; }

        protected abstract string GetSegmentName(int segment, int segmentCount);

        protected abstract Task<(DataTable data, int rowCount)> QuerySegmentAsync(int pageNumber, int pageSize);

        protected Stream BuildWorkbookStream(DataTable dataTable, Action<IXLWorksheet> formatWorksheet = null)
        {
            using (var wb = new XLWorkbook())            
            {
                var ws = wb.AddWorksheet(dataTable);
                formatWorksheet?.Invoke(ws);
                var ms = new MemoryStream();
                wb.SaveAs(ms);
                ms.Position = 0;
                return ms;
            }
        }

        protected virtual async Task InitializeAsync(TOperation operationId) => await Task.CompletedTask;

        private async Task BuildZipFileAsync(MemoryStream result, int segmentCount, TOperation operationId, Action<IXLWorksheet> formatWorksheet = null)
        {
            using (var zip = new ZipArchive(result, ZipArchiveMode.Create, true))
            {
                var segments = SegmentsPerZipEntry / segmentCount;
                var leftover = ((SegmentsPerZipEntry % segmentCount) > 0) ? 1 : 0;
                segments += leftover;

                int skip = 0;
                for (int segment = 0; segment < segments; segment++)
                {
                    var data = await _segmentStore.AssembleAsync(operationId, skip, SegmentsPerZipEntry);
                    
                    var entry = zip.CreateEntry(GetSegmentName(segment, segments));
                    using (var entryStream = entry.Open())
                    {
                        using (var wb = BuildWorkbookStream(data, formatWorksheet))
                        {
                            await wb.CopyToAsync(entryStream);
                        }                            
                    }

                    skip += SegmentsPerZipEntry;
                }
            }
        }
    }
}
