string SQLCommand = @”
UPDATE [table_requests]
SET status = ‘old’
WHERE [id] = @id AND (status = @status1 OR status = @status2)”;

using (SQLiteCommand command = new SQLiteCommand(SQLCommand, DB))
{
command.Parameters.AddWithValue(”@id”, Convert.ToInt32(DataRequests.Kod));
command.Parameters.AddWithValue(”@status1”, “В работе”);
command.Parameters.AddWithValue(”@status2”, “Новая”);

await command.ExecuteNonQueryAsync();

}
