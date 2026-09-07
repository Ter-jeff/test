using System.Diagnostics.CodeAnalysis;

namespace CommonLib.Results;

public sealed record Result<TResult>
{
    [MemberNotNullWhen(true, nameof(Value))]
    [MemberNotNullWhen(false, nameof(Error))]
    public bool Success { get; init; }

    public TResult? Value { get; init; }

    public string? Error { get; init; }

    public static Result<TResult> Ok(TResult tResult)
    {
        return new()
        {
            Success = true,
            Value = tResult
        };
    }

    public static Result<TResult> Fail(string errorMessage)
    {
        return new()
        {
            Success = false,
            Error = errorMessage
        };
    }
}

public sealed record Result<TResult, TError>
{
    [MemberNotNullWhen(true, nameof(Value))]
    [MemberNotNullWhen(false, nameof(Error))]
    public bool Success { get; init; }

    public TResult? Value { get; init; }

    public TError? Error { get; init; }

    public static Result<TResult, TError> Ok(TResult tResult)
    {
        return new()
        {
            Success = true,
            Value = tResult
        };
    }

    public static Result<TResult, TError> Fail(TError tError)
    {
        return new()
        {
            Success = false,
            Error = tError
        };
    }
}
