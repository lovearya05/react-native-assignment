import React, { useEffect, useState } from "react";

import {
  Button,
  Linking,
  Platform,
  StatusBar,
  StyleSheet,
  Text,
  TextInput,
  TouchableOpacity,
  View,
} from "react-native";

import * as FileSystem from "expo-file-system";
import ExcelJS from "exceljs";
import * as Sharing from "expo-sharing";
import { Buffer as NodeBuffer } from "buffer";
import AsyncStorage from "@react-native-async-storage/async-storage";
import { platform } from "os";
import { MIMEType } from "util";
import moment from "moment";

export default function App() {

  // state to save grid data
  // as required only 10 rows and 5 columns
  // i used 11 rows because 0th row i used for A,B,C,D, ...
  // i used 6 cols because 0th colunm used for 1,2,3 ...
  const [gridData, setGridData] = useState(
    Array.from({ length: 11 }, () => Array(6).fill(""))
  );

  const [projectName, setProjectName] = useState("");

    // getting data from local storage when first time open the app

    const loadDataFromAsyncStorage = async () => {
      try {
        AsyncStorage.getItem("sheetData").then((data) => {
          if (data) {
            const storedData = JSON.parse(data);
            setGridData(storedData);
          } else {
            console.log("no data"); /// remove
          }
        });
      } catch (error) {
        console.error("Error loading data from AsyncStorage:", error);
        return [];
      }
    };
  
    useEffect(() => {
      loadDataFromAsyncStorage();
    }, []);



    
  // saving data to local storage

  const saveDataToAsyncStorage = async (data) => {
    try {
      const str = JSON.stringify(gridData);
      // console.log(str)
      await AsyncStorage.setItem("sheetData", str).then((res) => {
        console.log(res, "data saved");
        // loadDataFromAsyncStorage()
      });
    } catch (error) {
      console.error("Error saving data to AsyncStorage:", error);
    }
  };

  useEffect(() => {
    saveDataToAsyncStorage(gridData);
  }, [gridData]);

  

  // code to generrate a sharable file URI

  const generateShareableExcel = async (data) => {
    let currentTime = moment().format("HH:mm:ss");
    const now = new Date();
    const fileName =  `${projectName}-${currentTime}.xlsx`;
    const fileUri = FileSystem.cacheDirectory + fileName;
    return new Promise((resolve, reject) => {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Me";
      workbook.created = now;
      workbook.modified = now;
      const worksheet = workbook.addWorksheet("My Sheet", {});

      data.forEach((rowData) => {
        worksheet.addRow(rowData);
      });

      workbook.xlsx.writeBuffer().then((buffer) => {
        const nodeBuffer = NodeBuffer.from(buffer);
        const bufferStr = nodeBuffer.toString("base64");
        FileSystem.writeAsStringAsync(fileUri, bufferStr, {
          encoding: FileSystem.EncodingType.Base64,
        }).then(() => {
          resolve(fileUri);
        });
      });
    });
  };

  
  // this function generate a downlodable file and save it to device storage

  const download = async () => {
    let currentTime = moment().format("HH:mm:ss");
    if (projectName === "" || projectName.trim().length === 0) {
      alert("Please fill the details");
      return;
    }
    let sheetArr = [];
    for (let ar of gridData) sheetArr.push(ar.slice(1, ar.length));
    const shareableExcelUri = await generateShareableExcel(
      sheetArr.slice(1, sheetArr.length)
    );

    // console.log(shareableExcelUri)

    const fileName = "abc";
    const localhost = Platform.OS === "android" ? "10.0.2.2" : "127.0.0.1";

    if (Platform.OS === "android") {
      const permission =
        await FileSystem.StorageAccessFramework.requestDirectoryPermissionsAsync();
      if (permission.granted) {
        const base64 = await FileSystem.readAsStringAsync(shareableExcelUri, {
          encoding: FileSystem.EncodingType.Base64,
        });
        await FileSystem.StorageAccessFramework.createFileAsync(
          permission.directoryUri,
          `${projectName}-${currentTime}.xlsx`,
          // "jaiHanumanJi.xlsx",
          "xlsx"
        )
          .then(async (uri) => {
            await FileSystem.writeAsStringAsync(uri, base64, {
              encoding: FileSystem.EncodingType.Base64,
            });
          })
          .catch((e) => console.log(e + "--er"));
      } else {
        Sharing.shareAsync(shareableExcelUri);
      }
    }
  };


  
  // this function generate a sharable file when a share button is pressed
  
  const shareExcel = async () => {
    if (projectName === "" || projectName.trim().length === 0) {
      alert("Please fill the details");
      return;
    }
    let sheetArr = [];
    for (let ar of gridData) sheetArr.push(ar.slice(1, ar.length));
    const shareableExcelUri = await generateShareableExcel(
      sheetArr.slice(1, sheetArr.length)
    );

    Sharing.shareAsync(shareableExcelUri, {
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      dialogTitle: "Your dialog title here",
      UTI: "com.microsoft.excel.xlsx",
    })
      .catch((error) => {
        console.error("Error", error);
      })
      .then(() => {
        console.log("Return from sharing dialog");
      });
  };



  // this is the function i used to clear all the data
  // just assigning a new array

  const clearAllData = ()=>{
    setGridData(Array.from({ length: 11 }, () => Array(6).fill("")))
  }



  // this is the main grid that is shown on the home screen

  const renderGrid = () => {
    const rows = [];
    let char = 65;
    for (let i = 0; i < 11; i++) {
      const row = [];
      for (let j = 0; j < 6; j++) {
        row.push(
          (i == 0 && j != 0) || (j == 0 && i > 0) ? (
            <TextInput
              key={`${i}-${j}`}
              style={{
                borderWidth: 0.5,
                width: "16.66%",
                textAlign: "center",
              }}
              onChangeText={(val) => {
                gridData[i][j] = val;
                setGridData([...gridData]);
              }}
              value={j == 0 ? i + "" : String.fromCharCode(char++)}
              editable={false}
            />
          ) 
          :
          (
            <TextInput
              key={`${i}-${j}`}
              style={{
                borderWidth: 0.5,
                width: "16.66%",
                textAlign: "center",
              }}
              onChangeText={(val) => {
                gridData[i][j] = val;
                setGridData([...gridData]);
              }}
              value={gridData[i] ? gridData[i][j] : ""}
            />
          )
        );
      }
      rows.push(
        <View key={i} style={{ flexDirection: "row", marginRight: 10 }}>
          {row}
        </View>
      );
    }
    return rows;
  };


  return (
    <>
      <StatusBar style="light" backgroundColor="#002F34" />
      <View
        style={styles.navbar}
      >
        <Text style={{ fontSize: 18, color: "white" }}>Lovepreet</Text>

        <View style={styles.buttonParent}>
          <TouchableOpacity
            style={styles.clearAll}
            onPress={clearAllData}
          >
            <Text style={styles.buttonText}>
              Clear All
            </Text>
          </TouchableOpacity>

          <TouchableOpacity
            onPress={download}
            style={styles.download}
          >
            <Text style={styles.buttonText}>
              Download
            </Text>
          </TouchableOpacity>

          <TouchableOpacity onPress={shareExcel}>
            <Text
              style={styles.share}
            >
              Share
            </Text>
          </TouchableOpacity>
        </View>
      </View>


      <View
        style={styles.gridMain}
      >
        <View
          style={styles.fileNameInput}
        >
          <TextInput
            style={ styles.textInputProjectName }
            placeholder="Enter file name"
            value={projectName}
            onChangeText={(text) => setProjectName(text)}
          />
        </View>
        {renderGrid()}
      </View>
    </>
  );


}


const styles = StyleSheet.create({
  navbar:{
    backgroundColor: "#002F34",
    width: "100%",
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    padding: 15,
  },
  buttonParent:{ 
    display: "flex", 
    flexDirection: "row" 
  },
  buttonText:{
     color: "white",
      fontSize: 16,
       fontWeight: "400"
  },
  clearAll:{
    borderRightWidth: 0.5,
    borderRightColor: "#EFEFEF",
    paddingRight: 10,
    marginRight: 15,
  },
  download : {
    borderRightWidth: 0.5,
    borderRightColor: "#EFEFEF",
    paddingRight: 10,
  },
  share:{
    marginLeft: 8,
    color: "white",
    fontSize: 16,
    fontWeight: "400",
  },
  textInputProjectName: {
    paddingLeft: 10,
    marginLeft: 5,
    borderColor: "grey",
    borderRadius: 5,
    borderWidth: 0.5,
    marginRight: 10,
    width: "97%",
  },

  gridMain:{
    marginTop: 15,
    justifyContent: "center",
    alignItems: "center",
    flexWrap: "wrap",
  },

  fileNameInput:{
    display: "flex",
    alignItems: "center",
    flexDirection: "row",
    marginBottom: 10,
    width: "100%",
    justifyContent: "flex-start",
  },

});