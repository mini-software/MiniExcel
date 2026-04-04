using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Comments;

public class CommentsRetrievalAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();

    [Fact]
    public async Task SheetWithCommentsAndNotesTestAsync()
    {
        var commentSet = await _excelImporter.RetrieveCommentsAsync(PathHelper.GetFile("xlsx/TestCommentsAndNotes.xlsx"), "sheet1");
        var (firstComment, secondComment) = (commentSet.Comments[0], commentSet.Comments[1]);
        
        Assert.Equal("sheet1", commentSet.SheetName, ignoreCase: true);
        Assert.Equal(2, commentSet.Comments.Count);
        
        Assert.Equal("B3", firstComment.ReferenceCell);
        Assert.Equal(new DateTime(2026, 3, 21, 12, 7, 24), firstComment.CreatedAt);
        Assert.Equal(new Guid("8d44beaf-9259-4d6a-8559-58427a76727b"), firstComment.Id);
        Assert.Equal("this is a comment", firstComment.Text);
        Assert.Equal(new Guid("cb8b42e9-e059-4d6b-b054-b1437d6cf7cd"), firstComment.Author?.Id);
        Assert.Equal("John Doe", firstComment.Author?.DisplayName);
        Assert.Equal("google-sheets", firstComment.Author?.ProviderId);
        Assert.Equal(2, firstComment.Replies.Count);
        Assert.False(firstComment.Resolved);

        Assert.Equal(new Guid("dfb1d4cd-7f1f-42ae-9f61-330f03f0b9ad"), firstComment.Replies[0].Id);
        Assert.Equal(new Guid("8d44beaf-9259-4d6a-8559-58427a76727b"), firstComment.Replies[0].ParentId);
        Assert.Equal(new DateTime(2026, 3, 21, 21, 17, 45), firstComment.Replies[0].CreatedAt);
        Assert.Equal("Mary Sue", firstComment.Replies[0].Author?.DisplayName);
        Assert.Equal("this is a reply", firstComment.Replies[0].Text);
        
        Assert.Equal(new Guid("d99bde2c-3df5-4300-a12e-2cc3b831c5dd"), firstComment.Replies[1].Id);
        Assert.Equal(new Guid("8d44beaf-9259-4d6a-8559-58427a76727b"), firstComment.Replies[1].ParentId);
        Assert.Equal(new DateTime(2026, 3, 21, 21, 20, 3), firstComment.Replies[1].CreatedAt);
        Assert.Equal("John Doe", firstComment.Replies[1].Author?.DisplayName);
        Assert.Equal("this is another reply", firstComment.Replies[1].Text);

        Assert.Empty(secondComment.Replies);
        Assert.Equal("E2", secondComment.ReferenceCell);
        Assert.Equal(new Guid("0fdf4b1e-0d47-4717-9dd5-c9fc731b0ad6"), secondComment.Id);
        Assert.Equal(new DateTime(2026, 3, 21, 21, 35, 17), secondComment.CreatedAt);
        Assert.Equal(new Guid("eaf7fda0-61e5-4210-9faa-da7028ea718a"), secondComment.Author?.Id);
        Assert.Equal("Mary Sue", secondComment.Author?.DisplayName);
        Assert.Equal("AD", secondComment.Author?.ProviderId);
        Assert.False(secondComment.Resolved);
        Assert.Equal("this is a separate comment", secondComment.Text);
        
        Assert.Equal(2, commentSet.Notes.Count);
        var (firstNote, secondNote) = (commentSet.Notes[0], commentSet.Notes[1]);
        
        Assert.Equal(new Guid("00000000-0006-0000-0000-000001000000"), firstNote.Id);
        Assert.Equal("D6", firstNote.ReferenceCell);
        Assert.Empty(firstNote.Author ?? "");
        Assert.Equal("this is a simple note", firstNote.Text);

        Assert.Equal(new Guid("4e01653b-66e0-48be-9390-2bddb28a7255"), secondNote.Id);
        Assert.Equal("G10", secondNote.ReferenceCell);
        Assert.Equal("local user", secondNote.Author);
        Assert.Equal("local user:\nthis is a note from someone else", secondNote.Text);
    }
    
    [Fact]
    public async Task SheetWithNotesAndCommentsWithoutRepliesTestAsync()
    {
        var commentSet = await _excelImporter.RetrieveCommentsAsync(PathHelper.GetFile("xlsx/TestCommentsAndNotes.xlsx"), "sheet2");
        var comment = commentSet.Comments[0];

        Assert.Equal("sheet2", commentSet.SheetName, ignoreCase: true);
        Assert.Single(commentSet.Comments);
        Assert.Empty(comment.Replies);
        Assert.Equal("A3", comment.ReferenceCell);
        Assert.Equal(new Guid("597d85de-079d-4129-8ebb-e6a9666c1c31"), comment.Id);
        Assert.Equal(new DateTime(2026, 3, 21, 12, 8, 22), comment.CreatedAt);
        Assert.Equal(new Guid("cb8b42e9-e059-4d6b-b054-b1437d6cf7cd"), comment.Author?.Id);
        Assert.Equal("John Doe", comment.Author?.DisplayName);
        Assert.Equal("google-sheets", comment.Author?.ProviderId);
        Assert.False(comment.Resolved);
        Assert.Equal("this is a comment on another sheet", comment.Text);

        Assert.Single(commentSet.Notes);
        var note = commentSet.Notes[0];

        Assert.Equal(new Guid("00000000-0006-0000-0100-000001000000"), note.Id);
        Assert.Equal("B11", note.ReferenceCell);
        Assert.Empty(commentSet.Notes[0].Author ?? "");
        Assert.Equal("this is a note on another sheet", note.Text);
    }

    [Fact]
    public async Task SheetWithoutNotesNorCommentsTestAsync()
    {
        var commentSet = await _excelImporter.RetrieveCommentsAsync(PathHelper.GetFile("xlsx/TestCommentsAndNotes.xlsx"), "sheet3");
        Assert.Equal("sheet3", commentSet.SheetName, ignoreCase: true);
        Assert.Empty(commentSet.Comments);
        Assert.Empty(commentSet.Notes);
    }
        
    [Fact]
    public async Task SheetWithResolvedThreadedCommentsTestAsync()
    {
        var commentSet = await _excelImporter.RetrieveCommentsAsync(PathHelper.GetFile("xlsx/TestCommentsAndNotes.xlsx"), "sheet4");
        var comment = commentSet.Comments[0];

        Assert.Single(commentSet.Comments);
        Assert.Equal("sheet4", commentSet.SheetName, ignoreCase: true);
        Assert.Equal("D2", comment.ReferenceCell);
        Assert.Equal(new DateTime(2026, 3, 21, 12, 34, 24), comment.CreatedAt);
        Assert.Equal(new Guid("cc210736-0fae-4525-aa57-df776a5548fa"), comment.Id);
        Assert.Equal("this thread will be resolved", comment.Text);
        Assert.Equal(new Guid("cb8b42e9-e059-4d6b-b054-b1437d6cf7cd"), comment.Author?.Id);
        Assert.Equal("John Doe", comment.Author?.DisplayName);
        Assert.Equal("google-sheets", comment.Author?.ProviderId);
        Assert.Single(comment.Replies);
        Assert.True(comment.Resolved);

        Assert.Equal(new Guid("f4863ec3-3a84-453a-88a7-ed634a96dd18"), comment.Replies[0].Id);
        Assert.Equal(new Guid("cc210736-0fae-4525-aa57-df776a5548fa"), comment.Replies[0].ParentId);
        Assert.Equal(new DateTime(2026, 3, 21, 21, 20, 55), comment.Replies[0].CreatedAt);
        Assert.Equal(new Guid("eaf7fda0-61e5-4210-9faa-da7028ea718a"), comment.Replies[0].Author?.Id);
        Assert.Equal("Mary Sue", comment.Replies[0].Author?.DisplayName);
        Assert.Equal("AD", comment.Replies[0].Author?.ProviderId);
        Assert.Equal("ok", comment.Replies[0].Text);
    }
}
