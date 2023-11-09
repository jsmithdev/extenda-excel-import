import { api, LightningElement } from 'lwc';
import { loadScript } from "lightning/platformResourceLoader";
import StaticRes from "@salesforce/resourceUrl/exceljs";
import ExtendaElement from 'c/extendaElement';

export default class ExtendaExcel extends ExtendaElement {

    @api
    newWorkbook(){
        return new ExcelJS.Workbook();
    }

    @api
    excelToObjects(file, columnsToFields, options = {}){

        return new Promise((resolve, reject) => {
            try{
                return this.excelToObjectsProcess(file, columnsToFields, options, resolve)
            }
            catch(e){
                reject(e);
            }
        });
    }

    connectedCallback() {
        loadScript(this, StaticRes)
        .then(() => {
            //console.log(ExcelJS);
        })
        .catch(error => {
            console.log(error);
        });
    }

    /**
     * Excel file to record objects
     * @param {File} file
     * @param {Object} columnsToFields - map of column names to field names; { 'column name': 'field name' } (tip: set 'field name' to 'skip' to not return as a field)
     * @param {Object} options
     * @param {Boolean} options.debug - log debug info
     * @param {Function} resolve - resolve function
     * @returns {Promise}
     */
    excelToObjectsProcess(file, columnsToFields, options, resolve){
        
        const Workbook = new ExcelJS.Workbook()
        const reader = new FileReader()
        
        reader.onload = async () => {

            const buffer = reader.result;

            const workbook = await Workbook.xlsx.load(buffer, {
                ignoreNodes: [
                    'picture',
                    'drawing',
                ],
            })
            
            const sheet = workbook.getWorksheet('main');
            
            const [columns, ...rows] = sheet._rows
                .filter(x => x.values?.length)
                .map(x => x.values)

            const fields = columns
                .map(x => columnsToFields[x])
                .filter(x => x)

            const records = rows.flatMap(row => {
                
                return row.reduce((acc,cur,i) => {
                    if(fields[i] && fields[i] !== 'skip'){
                        acc[fields[i]] = typeof cur === 'object' ? cur.result : cur;
                    }
                    return acc;
                }, {});
            });

            if(options.debug){
                console.log({
                    name: file.name,
                    tab: sheet.name,
                    fields,
                    records,
                    columns,
                    rows,
                })
            }

            resolve({
                name: file.name,
                tab: sheet.name,
                fields,
                records,
                columns,
                rows,
            })
        };

        reader.readAsArrayBuffer(file)
    }
}