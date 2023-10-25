namespace Coginov.GraphApi.Library.Helpers
{
    public static class ConstantHelper
    {
        public static readonly int DEFAULT_CHUNK_SIZE = 100 * 1024 * 1024; // 100 MB
        public static readonly int DEFAULT_RETRY_IN_SECONDS = 1;
        public static readonly int DEFAULT_RETRY_COUNT = 5;
    }
}