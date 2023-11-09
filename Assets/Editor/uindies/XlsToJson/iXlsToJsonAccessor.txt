using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public interface iXlsToJsonAccessor
{
    /// <summary>
    /// Initialize
    /// </summary>
    public void   Initialize();
    /// <summary>
    /// Check Serialized table Exists
    /// </summary>
    public bool   CheckSerializedTableExists();
}
