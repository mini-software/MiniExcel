using System;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal class MiniExcelTask
    {
#if NET462
        public static Task CompletedTask = Task.FromResult(0);
#else
        public static Task CompletedTask = Task.CompletedTask;
#endif

        public static Task FromException(Exception exception)
        {
#if NET462
            var tcs = new TaskCompletionSource<object>();
            tcs.SetException(exception);
            return tcs.Task;
#else
            return Task.FromException(exception);
#endif
        }

        public static Task<T> FromException<T>(Exception exception)
        {
#if NET462
            var tcs = new TaskCompletionSource<T>();
            tcs.SetException(exception);
            return tcs.Task;
#else
            return Task.FromException<T>(exception);
#endif
        }

        public static Task FromCanceled(CancellationToken cancellationToken)
        {
#if NET462
            var tcs = new TaskCompletionSource<object>();
            cancellationToken.Register(() => tcs.SetCanceled());
            return tcs.Task;
#else
            return Task.FromCanceled(cancellationToken);
#endif
        }

        public static Task<T> FromCanceled<T>(CancellationToken cancellationToken)
        {
#if NET462
            var tcs = new TaskCompletionSource<T>();
            cancellationToken.Register(() => tcs.SetCanceled());
            return tcs.Task;
#else
            return Task.FromCanceled<T>(cancellationToken);
#endif
        }
    }
}
