import React from 'react';
import './App.css';
import axios from 'axios';
import FileSaver from 'file-saver';

class App extends React.Component {
    constructor(props){
      super(props);
      this.state = {
        tableData: [
          {student: 'pavan','marks': '100'},
          {student: 'rahul','marks': '50'}
        ]
      }
      this.getHeader = this.getHeader.bind(this);
      this.getRowsData = this.getRowsData.bind(this);
      }
      getHeader = () => {
        var keys = Object.keys(this.state.tableData[0]);
        return keys.map((key, index)=>{
        return <th key={key}>{key.toUpperCase()}</th>
        })
        }
        convertToXls = () => {
          const axiosConfig = {
          headers : { responseType: 'blob'}
          };
          axios.get('http://localhost:8000/convertGridToExcel',axiosConfig).then((res) => {
            console.log('---in');
            let fileBlob = res;

      let blob = new Blob([fileBlob], {
        type: "application/vnd.ms-excel"
      });

      //use file saver npm package for saving blob to file
      FileSaver.saveAs(blob, `myFile.xls`);
          }).catch(err => {console.log('err')})
        }
        getRowsData = () => {
          var items = this.state.tableData;
          var keys = Object.keys(this.state.tableData[0]);
          return items.map((row, index)=>{
          return <tr key={index}><RenderRow key={index} data={row} keys={keys}/></tr>
          })
          }
  render(){
  return (
    <div>
 <table>
 <thead>
 <tr>{this.getHeader()}</tr>
 </thead>
 <tbody>
 {this.getRowsData()}
 </tbody>
 </table>
 <div><button onClick = {this.convertToXls}>ConvertToXls</button></div>
 </div>
  );
  }
}
const RenderRow = (props) =>{
  return props.keys.map((key, index)=>{
  return <td key={props.data[key]}>{props.data[key]}</td>
  })
 }
export default App;
