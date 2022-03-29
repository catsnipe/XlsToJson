using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class UsageAccess : MonoBehaviour
{
    void Start()
    {
        // Json Data
        Debug.Log("--- Json Data ---");

        var jsontext = Resources.Load<TextAsset>("Data_Character");
        if (jsontext == null)
        {
            Debug.LogWarning("require Tools/XlsToJson/[Create]JsonData");
        }
        else
        {
            Character.SetTable(jsontext);
            Character.Rows.ForEach((row) => Debug.Log($"{row.ID} {row.CharacterName} {row.Sex}"));
        }

        // Scriptable Object
        Debug.Log("--- Scriptable Object ---");

        var obj = Resources.Load<Object>("Data_Character");
        if (obj == null)
        {
            Debug.LogWarning("require Tools/XlsToJson/[Create]Scriptable Object");
        }
        else
        {
            Character.SetTable(obj);
            Character.Rows.ForEach((row) => Debug.Log($"{row.ID} {row.CharacterName} {row.Sex}"));
        }

    }
}
