namespace MiniExcelLib.OpenXml.Models;

public class CommentResultSet
{
    public IReadOnlyList<ThreadedComment> Comments { get; internal set; } = [];
    public IReadOnlyList<NoteComment> Notes { get; internal set; } = [];
}

public class ThreadedComment
{
    public Guid Id { get; internal set; }
    public string ReferenceCell { get; internal set; } = null!;
    public Author? Author { get; internal set; }
    public bool Active { get; internal set; }
    public string? FirstMessage { get; internal set; }
    public DateTime CreatedAt { get; internal set; }

    internal List<ThreadedCommentReply> ThreadedComments = [];
    public IReadOnlyList<ThreadedCommentReply> Replies => ThreadedComments;
}

public class ThreadedCommentReply
{
    public Guid Id { get; internal set; }
    public Guid? ParentId { get; internal set; }
    public Author? Author { get; internal set; }
    public DateTime ReplyTime { get; internal set; }
    public string? Text { get; internal set; }
}

public class NoteComment
{
    public Guid Id { get; internal set; }
    public string? ReferenceCell { get; internal set; }
    public string? Author { get; internal set; }
    public string? Text { get; internal set; }
}

public class Author
{
    public Guid Id { get; internal set; }
    public string DisplayName { get; internal set; } = "";
    public string? ProviderId { get; internal set; }
}
