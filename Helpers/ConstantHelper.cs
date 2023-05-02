namespace Coginov.GraphApi.Library.Helpers
{
    public static class ConstantHelper
    {
        public static readonly long DEFAULT_CHUNK_SIZE = 1024 * 1024 * 1024; // 1 GB
        public static readonly int DEFAULT_RETRY_IN_SECONDS = 1;
        public static readonly int DEFAULT_RETRY_COUNT = 5;
    }
}