using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class UsageTable : MonoBehaviour
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
            var data = Class_Character.CreateInstance();
            data.SetRowsByJson(jsontext.ToString());

            foreach (var row in data.Rows)
            {
                Debug.Log($"{row.ID} {row.CharacterName} {row.Sex}");
            }
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
            var data2 = obj as Class_Character;

            foreach (var row in data2.Rows)
            {
                Debug.Log($"{row.ID} {row.CharacterName} {row.Sex}");
            }
        }

    }
}
