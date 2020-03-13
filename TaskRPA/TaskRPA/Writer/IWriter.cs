using System.Collections.Generic;

namespace TaskRPA.Writer
{
    interface IWriter<T>
    {
        void Write(List<T> list);
    }
}
