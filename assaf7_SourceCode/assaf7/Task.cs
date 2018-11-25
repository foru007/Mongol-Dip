using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace assaf7
{
    [Serializable]
    public class Task:IComparable<Task>
    {
        public static List<Task> CurrentQueue;
        public string Name;
        public DateTime Time;
        public string Notes;
        public DateTime now;

        public Task(string name,DateTime time,string notes)
        {
            Name = name;
            Time = time;
            Notes = notes;
            now = DateTime.Now;
        }
        public Task(string name, DateTime time)
            : this(name, time, "")
        { }
        public Task(string name)
        {
            Name = name;
        }
        #region IComparable<Task> Members

        public int CompareTo(Task other)
        {
            return Time.CompareTo(other.Time);
        }

        #endregion
    }
}
