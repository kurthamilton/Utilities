using System;
using System.Collections;
using System.Collections.Generic;

namespace Utilities.Office
{
    public interface IOfficeCollection<T> where T : BaseOffice
    {        
        // Implement IEnumerable
        IEnumerator<T> GetEnumerator();

        // PUBLIC METHODS        
        T this[int index] { get; }

        bool Contains(int index);

        T Insert(int index);
        void Delete(int index);
    }
}
