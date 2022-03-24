using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Text;

namespace MiniExcelLibs
{
    public class DbList : IList<string>, IDisposable
    {
        private SQLiteConnection _conn;
        private SQLiteCommand _cmd;
        private string _name;
        private const string _tableName = "sharedStrings";

        public DbList(string name)
        {
            _name = name;
            _conn = new SQLiteConnection($"Data Source={name}.db;Version=3;");
            _conn.Open();
            _cmd = _conn.CreateCommand();

            CreateTable();
        }

        private void CreateTable()
        {
            Clear();
            _cmd.CommandText = $@"
CREATE TABLE {_tableName} (name TEXT, `index` INTEGER);

CREATE UNIQUE INDEX idx_index
ON sharedStrings (
  `index`
);

CREATE INDEX idx_name
ON sharedStrings (
  name
);";
            _cmd.ExecuteNonQuery();
        }

        public IEnumerator<string> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        public void Add(string item)
        {
            var maxIndex = GetMaxIndex();
            _cmd.CommandText = $"INSERT INTO {_tableName}(name, `index`) VALUES ('{item}', {maxIndex + 1})";
            _cmd.ExecuteNonQuery();
        }

        private long GetMaxIndex()
        {
            _cmd.CommandText = $"SELECT MAX(`index`) FROM {_tableName}";
            var result = _cmd.ExecuteScalar();
            if (result == DBNull.Value)
                return -1;

            return (long)result;
        }

        public void Clear()
        {
            _cmd.CommandText = $"DROP TABLE IF EXISTS {_tableName}";
            _cmd.ExecuteNonQuery();
        }

        public bool Contains(string item)
        {
            _cmd.CommandText = $"SELECT * FROM {_tableName} WHERE name = '{item}'";
            return _cmd.ExecuteScalar() != null;
        }

        public void CopyTo(string[] array, int arrayIndex)
        {
            throw new System.NotImplementedException();
        }

        public void AddRange(List<string> array)
        {
            var maxIndex = GetMaxIndex();

            var cmdTxt = new StringBuilder();

            cmdTxt.Append($"INSERT INTO {_tableName}(name, `index`) VALUES");
            for (var i = 0; i < array.Count; i++)
            {
                var item = array[i];
                cmdTxt.Append($"('{item}', {maxIndex + i + 1})");
                cmdTxt.Append(i != array.Count - 1 ? ',' : ';');
            }

            _cmd.CommandText = cmdTxt.ToString();
            _cmd.ExecuteNonQuery();
        }

        public bool Remove(string item)
        {
            _cmd.CommandText = $"DELETE FROM {_tableName} WHERE name = '{item}'";
            return _cmd.ExecuteNonQuery() > 0;
        }

        public int Count
        {
            get
            {
                _cmd.CommandText = "SELECT COUNT(*) FROM " + _tableName;
                return Convert.ToInt32(_cmd.ExecuteScalar());
            }
        }

        public bool IsReadOnly { get; }

        public int IndexOf(string item)
        {
            _cmd.CommandText = $"SELECT `index` FROM {_tableName} WHERE name = '{item}'";
            return (int)_cmd.ExecuteScalar();
        }

        public void Insert(int index, string item)
        {
            _cmd.CommandText = $"UPDATE {_tableName} SET `index` = `index` + 1 WHERE index >= {index}";
            _cmd.ExecuteNonQuery();
            _cmd.CommandText = $"INSERT INTO {_tableName}(name, `index`) VALUES ('{item}', {index})";
            _cmd.ExecuteNonQuery();
        }

        public void RemoveAt(int index)
        {
            _cmd.CommandText = $"DELETE FROM {_tableName} WHERE `index` = {index}";
            _cmd.ExecuteNonQuery();
        }

        public string this[int index]
        {
            get
            {
                _cmd.CommandText = $"SELECT name FROM {_tableName} WHERE `index` = {index}";
                return (string)_cmd.ExecuteScalar();
            }
            set
            {
                _cmd.CommandText = $"UPDATE {_tableName} SET name = '{value}' WHERE `index` = {index}";
                _cmd.ExecuteNonQuery();
            }
        }

        public void Dispose()
        {
            if (_cmd != null)
            {
                _cmd.Dispose();
                _cmd = null;
            }

            if (_conn != null)
            {
                _conn.Dispose();
                _conn = null;
            }

            File.Delete($"{_name}.db");
        }
    }
}