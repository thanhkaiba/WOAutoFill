
<script>

import * as XLSX from 'xlsx';
import axios from 'axios'
import * as Swal from 'sweetalert2'
import Loading from 'vue-loading-overlay';

import 'vue-loading-overlay/dist/css/index.css';
export default {
  data() {
    return {
      form: {
        formName: "Auto Fill Data",
        url: "",
        style: ""
      },

      fillPkgFirst: false,
      list: [],
      listFail: [],
      submitable: true,
      isLoading: false,
      sheet: "Demand",
      showSubmitFeedback: false
    }
  },
  components: {
    Loading,
  },
  methods: {
    async getSize(style, color, size, SizeDesc) {
      let data = `Style_Cd=${style}&Color_Cd=${color}&Attribute_Cd=------&Size_Cd=${size}`

      let config = {
        method: 'post',
        maxBodyLength: Infinity,
        url: 'http://wsscplanprd05/ISS/Order/GetSkuSizes',
        headers: {
          'Accept': '*/*',
          'Accept-Language': 'en-US,en;q=0.9',
          'Connection': 'keep-alive',
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'Cookie': 'menustate=false',
          'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36',
          'X-Requested-With': 'XMLHttpRequest'
        },
        data: data
      };

      try {
        let response = await axios.request(config);
        for (let i = 0; i < response.data.length; i++) {
          if (response.data[i]["SizeDesc"] == SizeDesc) {
            return response.data[i]["Size"];
          }
        }

      } catch (e) {

      }
      return "";

    },
    ExcelDateToJSDate(serial) {
      let utc_days = Math.floor(serial - 25569);
      let utc_value = utc_days * 86400;
      let date_info = new Date(utc_value * 1000);

      let fractional_day = serial - Math.floor(serial) + 0.0000001;

      let total_seconds = Math.floor(86400 * fractional_day);

      let seconds = total_seconds % 60;

      total_seconds -= seconds;

      let minutes = Math.floor(total_seconds / 60) % 60;

      const utc = Date.UTC(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), 5, minutes, seconds);
      return new Date(utc);
    },
    convertDate(dateStr) {
      if (dateStr == null || dateStr.length === 0) {
        return dateStr;
      }
      const milliseconds = parseInt(dateStr.match(/\/Date\((\d+)\)\//)[1]);
      const date = new Date(milliseconds);
      return date.toISOString();
    },
    convertDateXLSX(dateStr) {
      if (dateStr == null || dateStr.length === 0) {
        return "";
      }
      const configDate = new Date(dateStr);
      const utc = Date.UTC(configDate.getFullYear(), configDate.getMonth(), configDate.getDate(), 4, 0, 0, 0);
      return (new Date(utc)).toISOString();
    },
    async fillAll() {
      const xlsxfile = this.$refs.file.files[0];

      if (xlsxfile == null) {
        Swal.fire(
          'The XLSX file?',
          'Please xlsx file!',
          'question'
        );
        return;
      }

      let ws;
      try {
        const workbook = XLSX.read(await xlsxfile.arrayBuffer(), { type: 'binary' });
        ws = workbook.Sheets[this.sheet];
      } catch (e) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'Something went wrong!' + e.message,
          footer: '<a href="">Why do I have this issue?</a>'
        })
        return;
      }


      let styleList = new Set();
      let range = XLSX.utils.decode_range(ws['!ref']);
      for (let i = 0; i <= range.e.r + 1; i++) {
        if (ws["D" + i] != null && ws["D" + i].v.length > 0) {
          const style = ws["D" + i].v;
          styleList.add(style);

        }
      }
      console.log(styleList.size);

      for (let s of styleList) {
        try {
          console.log(s);
          this.form.style = s;
          await this.fakeSubmit();
          await this.submitdate();
          console.log("DONE");
        } catch (e) {
          Swal.fire({
            icon: 'error',
            title: 'Fill fail with style ' + this.form.style,
            text: e,
            footer: '<a href="">Why do I have this issue?</a>'
          });
        } finally {
          this.form.style = "";
        }
      }


    },
    async fakeSubmit() {
      this.listFail = [];
      const xlsxfile = this.$refs.file.files[0];

      if (this.form.style.length === 0) {
        Swal.fire(
          'The Style?',
          'Please style code!',
          'question'
        );
        return;
      }

      if (xlsxfile == null) {
        Swal.fire(
          'The XLSX file?',
          'Please xlsx file!',
          'question'
        );
        return;
      }

      let ws;
      try {
        const workbook = XLSX.read(await xlsxfile.arrayBuffer(), { type: 'binary' });
        ws = workbook.Sheets[this.sheet];
      } catch (e) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'Something went wrong!' + e.message,
          footer: '<a href="">Why do I have this issue?</a>'
        })
        return;
      }

      this.list = [];
      this.isLoading = true;


      try {
        let range = XLSX.utils.decode_range(ws['!ref']);
        for (let i = 0; i <= range.e.r + 1; i++) {
          if (ws["D" + i] != null && ws["D" + i].v === this.form.style) {
            if (ws["Q" + i] != null && ws["Q" + i].v != null) {

              if (ws["Y" + i] == null || ws["Y" + i].v == null || ws["Y" + i].v.length === 0) {
                throw "Missing Due Date at row " + i;
              }
              let q = "" + ws["Q" + i].v;

              if (q.indexOf("+") > 0) {

                q.split("+").forEach(e => {
                  if (!isNaN(e) || e.indexOf("*") > 0) {
                    if (e.indexOf("*") > 0) {
                      const multi = e.split("*");
                      for (let j = 0; j < +multi[1]; j++) {
                        this.list.push({
                          "style": ws["D" + i].v,
                          "color": ws["E" + i].v.trim(),
                          "size": ws["G" + i].v,
                          "quatity": +multi[0],
                          "pkg": ws["F" + i].v,
                          "dc": ws["K" + i].v.trim(),
                          "priority": ws["X" + i].v,
                          "revision": ws["I" + i] != null && ("" + ws["I" + i].v).trim().length > 0 ? +ws["I" + i].v : 0,
                          "duedate": this.ExcelDateToJSDate(ws["Y" + i].v),
                        });
                      }
                    } else {
                      this.list.push({
                        "style": ws["D" + i].v,
                        "color": ws["E" + i].v.trim(),
                        "size": ws["G" + i].v,
                        "quatity": +e,
                        "pkg": ws["F" + i].v,
                        "dc": ws["K" + i].v.trim(),
                        "priority": ws["X" + i].v,
                        "revision": ws["I" + i] != null && ("" + ws["I" + i].v).trim().length > 0 ? +ws["I" + i].v : 0,
                        "duedate": this.ExcelDateToJSDate(ws["Y" + i].v),
                      });
                    }
                  }

                });
              } else {

                if (q.indexOf("*") > 0) {
                  const multi = q.split("*");
                  for (let j = 0; j < +multi[1]; j++) {
                    this.list.push({
                      "style": ws["D" + i].v,
                      "color": ws["E" + i].v.trim(),
                      "size": ws["G" + i].v,
                      "quatity": +multi[0],
                      "pkg": ws["F" + i].v,
                      "dc": ws["K" + i].v.trim(),
                      "priority": ws["X" + i].v,
                      "revision": ws["I" + i] != null && ("" + ws["I" + i].v).trim().length > 0 ? +ws["I" + i].v : 0,
                      "duedate": this.ExcelDateToJSDate(ws["Y" + i].v),
                    });
                  }
                } else {
                  this.list.push({
                    "style": ws["D" + i].v,
                    "color": ws["E" + i].v.trim(),
                    "size": ws["G" + i].v,
                    "quatity": +q,
                    "pkg": ws["F" + i].v,
                    "dc": ws["K" + i].v.trim(),
                    "priority": ws["X" + i].v,
                    "revision": ws["I" + i] != null && ("" + ws["I" + i].v).trim().length > 0 ? +ws["I" + i].v : 0,
                    "duedate": this.ExcelDateToJSDate(ws["Y" + i].v),
                  });
                }
              }

            }

          }
        }
      } catch (e) {
        this.isLoading = false;
        if (this.list.length === 0) {
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: e,
            footer: '<a href="">Why do I have this issue?</a>'
          });
          return;
        }

      }

      this.isLoading = false;
      if (this.list.length === 0) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'No data found for style ' + this.form.style + ' !',
          footer: '<a href="">Why do I have this issue?</a>'
        });
        return;
      }

      this.submitable = true;
      this.showSubmitFeedback = true;



    },
    async filldata(searchData, list) {
      let lockedItem = [];

      for (let i = 0; i < searchData["Total"]; i++) {

        if (searchData["Data"][i]["OrderStatusDesc"] === "Locked") {
          lockedItem.push(searchData["Data"][i]);
        }
      }


      if (lockedItem.length < list.length) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'Not enough locked item for style ' + this.form.style + ' !',
          footer: '<a href="">Why do I have this issue?</a>'
        });
      } else {
        const editedItem = [];
        for (let i = 0; i < list.length; i++) {

          let locked = null;
          if (list[i]["SuperOrder"] == null || list[i]["SuperOrder"].length == 0) {
            locked = lockedItem[i];
          } else {
            for (let k = 0; k < lockedItem.length; k++) {
              if (lockedItem[k]["SuperOrder"] === list[i]["SuperOrder"]) {
                locked = lockedItem[k];
                break;
              }
            }
          }
          try {
            const Size = await this.getSize(locked.Style, list[i].color, locked["Size"], list[i].size)


            locked["CCurrDueDate"] = this.convertDate(locked["CCurrDueDate"]);
            locked["CurrDueDate"] = this.convertDate(locked["CurrDueDate"]);

            locked["StartDate"] = this.convertDate(locked["StartDate"]);
            locked["CStartDate"] = this.convertDate(locked["CStartDate"]);

            locked["EarliestStartDate"] = this.convertDate(locked["EarliestStartDate"]);
            locked["DemandDate"] = this.convertDate(locked["DemandDate"]);
            let Cloned = JSON.parse(JSON.stringify(locked));



            Cloned["idField"] = "Id";
            Cloned["_defaultId"] = 0;

            if (locked["CCurrDueDate"] == null) {
              locked["CCurrDueDate"] = locked["CurrDueDate"]
            }

            if (locked["CStartDate"] == null) {
              locked["CStartDate"] = locked["StartDate"]
            }

            locked["IsEdited"] = true;
            locked["Cloned"] = Cloned;
            locked["IsFieldChange"] = true;
            locked["Completed"] = false;
            locked["TotalDozens"] = list[i].quatity;
            locked["SizeShortDes"] = list[i].size;
            locked["Size"] = Size;

            locked["ExpeditePriority"] = list[i].priority;
            locked["DcLoc"] = list[i].dc;
            locked["Style"] = list[i].pkg;
            locked["Revision"] = list[i].revision;
            locked["Color"] = list[i].color;
            editedItem.push({
              item: locked,
              origin: list[i],
            });
          } catch (e) {
            Swal.fire({
              icon: 'error',
              title: 'Oops...',
              text: e,
              footer: '<a href="">Why do I have this issue?</a>'
            });
            return;
          }



        }


        const config = {
          method: 'post',
          headers: {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json; charset=UTF-8',
            'Cookie': 'menustate=false',
            'Origin': 'http://wsscplanprd05',
            'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
          }
        };
        this.listFail = [];
        try {
          for (let i = 0; i < editedItem.length; i++) {
            axios.post('http://wsscplanprd05/ISS/Order/SaveWOMdata', {
              "data": [editedItem[i].item],
              "mode": "Recalc"
            }, config).then(res => {
              console.log(res.data);
              if (res.data["Status"] === false) {
                this.listFail.push(editedItem[i].origin);
              }
            }).catch(e => {
              this.listFail.push(editedItem[i].origin);
            });

          }


          Swal.fire({
            position: 'top-end',
            icon: 'success',
            title: 'Your work has been saved',
            showConfirmButton: false,
            timer: 1500
          }).then(e => {
            this.form.style = "";
            this.list = [];
            this.showSubmitFeedback = true;
            this.submitable = true;
          })
        } catch (error) {
          // Handle errors
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            footer: '<a href="">Why do I have this issue?</a>'
          }).then(e => {
            this.submitable = true;
          });
        }
      }

    },
    async filldataWithOrder(searchData, list) {
      let lockedItem = [];

      for (let i = 0; i < searchData["Total"]; i++) {

        if (searchData["Data"][i]["OrderStatusDesc"] === "Locked") {
          lockedItem.push(searchData["Data"][i]);
        }
      }


      if (lockedItem.length < list.length) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'Not enough locked item for style ' + this.form.style + ' !',
          footer: '<a href="">Why do I have this issue?</a>'
        });
      } else {
        const editedItem = [];
        for (let i = 0; i < list.length; i++) {

          let locked = null;
          if (list[i]["SuperOrder"] == null || list[i]["SuperOrder"].length == 0) {
            locked = lockedItem[i];
          } else {
            for (let k = 0; k < lockedItem.length; k++) {
              if (lockedItem[k]["SuperOrder"] === list[i]["SuperOrder"]) {
                locked = lockedItem[k];
                break;
              }
            }
          }
          try {
            locked = JSON.parse(JSON.stringify(locked));

            locked["CCurrDueDate"] = this.convertDate(locked["CCurrDueDate"]);
            locked["CurrDueDate"] = this.convertDate(locked["CurrDueDate"]);

            locked["StartDate"] = this.convertDate(locked["StartDate"]);
            locked["CStartDate"] = this.convertDate(locked["CStartDate"]);

            locked["EarliestStartDate"] = this.convertDate(locked["EarliestStartDate"]);
            locked["DemandDate"] = this.convertDate(locked["DemandDate"]);
            let Cloned = JSON.parse(JSON.stringify(locked));



            Cloned["idField"] = "Id";
            Cloned["_defaultId"] = 0;

            if (locked["CCurrDueDate"] == null) {
              locked["CCurrDueDate"] = locked["CurrDueDate"]
            }

            if (locked["CStartDate"] == null) {
              locked["CStartDate"] = locked["StartDate"]
            }

            locked["IsEdited"] = true;
            locked["Cloned"] = Cloned;
            locked["IsFieldChange"] = true;
            locked["Completed"] = false;

            locked["Style"] = list[i].pkg;
            locked["Revision"] = list[i].revision;
            editedItem.push({
              item: locked,
              origin: list[i],
            });
          } catch (e) {
            Swal.fire({
              icon: 'error',
              title: 'Oops...',
              text: e,
              footer: '<a href="">Why do I have this issue?</a>'
            });
            return;
          }
        }


        const config = {
          method: 'post',
          headers: {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json; charset=UTF-8',
            'Cookie': 'menustate=false',
            'Origin': 'http://wsscplanprd05',
            'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
          }
        };
        try {
          for (let i = 0; i < editedItem.length; i++) {
            axios.post('http://wsscplanprd05/ISS/Order/SaveWOMdata', {
              "data": [editedItem[i].item],
              "mode": "Recalc"
            }, config).then(res => {
              if (res.data["Status"] === false) {
                console.log("fill with order fail");
                console.log(res.data);
              }
            }).catch(e => {
            });

          }

          await this.filldata(searchData, list)

        } catch (error) {
          // Handle errors
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            footer: '<a href="">Why do I have this issue?</a>'
          }).then(e => {
            this.submitable = true;
          });
        }
      }

    },
    async filldate(searchData, list) {
      let lockedItem = [];

      for (let i = 0; i < searchData["Total"]; i++) {

        if (searchData["Data"][i]["OrderStatusDesc"] === "Locked") {
          lockedItem.push(searchData["Data"][i]);
        }
      }


      if (lockedItem.length < list.length) {
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'Not enough locked item for style ' + this.form.style + ' !',
          footer: '<a href="">Why do I have this issue?</a>'
        });


      } else {
        let editedItem = [];
        for (let i = 0; i < list.length; i++) {
          list[i]["SuperOrder"] = lockedItem[i]["SuperOrder"];
          lockedItem[i]["CCurrDueDate"] = this.convertDate(lockedItem[i]["CCurrDueDate"]);
          lockedItem[i]["CurrDueDate"] = this.convertDate(lockedItem[i]["CurrDueDate"]);

          lockedItem[i]["StartDate"] = this.convertDate(lockedItem[i]["StartDate"]);
          lockedItem[i]["CStartDate"] = this.convertDate(lockedItem[i]["CStartDate"]);

          lockedItem[i]["EarliestStartDate"] = this.convertDate(lockedItem[i]["EarliestStartDate"]);
          lockedItem[i]["DemandDate"] = this.convertDate(lockedItem[i]["DemandDate"]);
          const Cloned = JSON.parse(JSON.stringify(lockedItem[i]));

          Cloned["idField"] = "Id";
          Cloned["_defaultId"] = 0;


          if (lockedItem[i]["CCurrDueDate"] == null) {
            lockedItem[i]["CCurrDueDate"] = lockedItem[i]["CurrDueDate"]
          }

          if (lockedItem[i]["CStartDate"] == null) {
            lockedItem[i]["CStartDate"] = lockedItem[i]["StartDate"]
          }


          lockedItem[i]["IsEdited"] = true;
          lockedItem[i]["IsFieldChange"] = true;
          lockedItem[i]["Completed"] = false;
          lockedItem[i]["Cloned"] = Cloned;
          lockedItem[i]["CurrDueDate"] = list[i].duedate.toISOString();

          editedItem.push({
            item: lockedItem[i],
            origin: list[i],
          });

        }


        const config = {
          method: 'post',
          headers: {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json; charset=UTF-8',
            'Cookie': 'menustate=false',
            'Origin': 'http://wsscplanprd05',
            'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
          }
        };
        this.listFail = [];

        for (let i = 0; i < editedItem.length; i++) {
          axios.post('http://wsscplanprd05/ISS/Order/SaveWOMdata', {
            "data": [editedItem[i].item],
            "mode": "EditPFSUngroup"
          }, config).then(res => {

            if (res.data["Status"] === false) {
              this.listFail.push(editedItem[i].origin);
            }
          }).catch(e => {
            this.listFail.push(editedItem[i].origin);
          });

        }
      }


    },
    async submit() {
      Swal.fire({
        title: 'Are you sure?',
        text: "You won't be able to revert this!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: "Yes, I'm sure!"
      }).then(async (result) => {
        if (result.isConfirmed) {
          await this.startFillData();
        }
      })




    },
    async startFillData() {
      this.submitable = false;
      let data = 'sort=&group=&filter=&SuperOrder=&StyleType=Selling+Style&SStyle=' + this.form.style + '&SColor=&SAttribute=&SSize=&DC=&Rev=&MfgPathId=95&Rule=&GroupId=&MFGPlant=&CylinderSize=&DyeBle=&TextileGroup=&Alt=&Machine=&Yarn=&DueDate=Earliest+Start&Week_input=Current+%2B+Prior+Week&Week=Current+%2B+Prior+Week&MoreWeeks_input=52&MoreWeeks=52&BOMMismatches=false&Fabric=&SuggestedLots=true&SpillOver=true&LockedLots=true&ReleasedLotsThisWeek=true&CustomerOrders=true&Events=true&MaxBuild=true&TILs=true&Forecast=false&StockTarget=true&Planner=&WorkCenter=&CapacityGroup=&CorpDiv=&BusinessUnit=&Src=A';

      let config = {
        method: 'post',
        url: 'http://wsscplanprd05/ISS/Order/WOManagement',
        headers: {
          'Accept': '*/*',
          'Accept-Language': 'en-US,en;q=0.9',
          'Connection': 'keep-alive',
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'Cookie': 'menustate=false',
          'Origin': 'http://wsscplanprd05',
          'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
          'X-Requested-With': 'XMLHttpRequest',
          'Access-Control-Allow-Origin': '*'
        },
        data: data
      };


      axios.request(config)
        .then(async (response) => {
          if (this.form.fillPkgFirst) {
            await this.filldataWithOrder(response.data, this.list);
          } else {
            await this.filldata(response.data, this.list);
          }
        })
        .catch((error) => {
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: error.message,
            footer: '<a href="">Why do I have this issue?</a>'
          });
        });

    },
    async submitdate() {
      this.isLoading = true;
      this.submitable = false;
      let data = 'sort=&group=&filter=&SuperOrder=&StyleType=Selling+Style&SStyle=' + this.form.style + '&SColor=&SAttribute=&SSize=&DC=&Rev=&MfgPathId=95&Rule=&GroupId=&MFGPlant=&CylinderSize=&DyeBle=&TextileGroup=&Alt=&Machine=&Yarn=&DueDate=Earliest+Start&Week_input=Current+%2B+Prior+Week&Week=Current+%2B+Prior+Week&MoreWeeks_input=52&MoreWeeks=52&BOMMismatches=false&Fabric=&SuggestedLots=true&SpillOver=true&LockedLots=true&ReleasedLotsThisWeek=true&CustomerOrders=true&Events=true&MaxBuild=true&TILs=true&Forecast=false&StockTarget=true&Planner=&WorkCenter=&CapacityGroup=&CorpDiv=&BusinessUnit=&Src=A';

      let config = {
        method: 'post',
        url: 'http://wsscplanprd05/ISS/Order/WOManagement',
        headers: {
          'Accept': '*/*',
          'Accept-Language': 'en-US,en;q=0.9',
          'Connection': 'keep-alive',
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'Cookie': 'menustate=false',
          'Origin': 'http://wsscplanprd05',
          'Referer': 'http://wsscplanprd05/ISS/Order/WOManagement',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
          'X-Requested-With': 'XMLHttpRequest',
          'Access-Control-Allow-Origin': '*'
        },
        data: data
      };


      axios.request(config)
        .then(async (response) => {
          await this.filldate(response.data, this.list);
          await this.startFillData();
        })
        .catch((error) => {
          Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: error.message,
            footer: '<a href="">Why do I have this issue?</a>'
          });
        }).finally(() => {
          this.isLoading = false;
        });




    }
  }
}
</script>

<template>
  <header>
    <link href="https://unpkg.com/nprogress@0.2.0/nprogress.css" rel="stylesheet" />
  </header>

  <div class="columns">
    <form class="example">
      <div class="row">
        <img src="./assets/duckdeveloper.png" alt="App Logo" width="100" height="100">
        <h3 style="color: blue;"><b>Note</b></h3>
        <ol>
          <li style="color: red;">Ensure that the cells in the excel file do not contain extra spaces.</li>
          <li style="color: red;">Create records equal to the number of records found in the excel file.</li>
          <li style="color: red;">When creating a new record, create the type of color code that has the most Sizes.</li>
        </ol>
      </div>
      <div class="row">
        <label for="style_code" class="label">Style Code</label>
        <div class="control">
          <input id="style_code" class="input" type="text" v-model="form.style" />
        </div>
      </div>

      <div class="row">
        <label for="xlsx_file" class="label">Excel file</label>
        <div class="control">
          <input id="xlsx_file" type="file" ref="file">
        </div>
      </div>

      <div class="custom-checkbox">
        <input type="checkbox" id="checkbox" v-model="form.fillPkgFirst" />
        <label for="checkbox">Fill PKG and Revision First</label>
      </div>


      <button type="submit" @click.prevent="fakeSubmit"><i class="fa fa-search">Search</i></button>
      <button type="submit" @click.prevent="fillAll"><i class="fa fa-search">Fill All(Beta)</i></button>

    </form>

    <div v-show="showSubmitFeedback && listFail.length > 0">
      <b style="color: red;">Items below was not success, need manual edit in website!!!</b>
      <table class="styled-table2">
        <thead>
          <tr>
            <th class="text-left">
              Style
            </th>
            <th class="text-left">
              Color
            </th>
            <th class="text-left">
              Size
            </th>
            <th class="text-left">
              Quatity
            </th>
            <th class="text-left">
              PKG Style
            </th>
            <th class="text-left">
              DC
            </th>
            <th class="text-left">
              Priority
            </th>
            <th class="text-left">
              Revision
            </th>
            <th class="text-left">
              Sew Due
            </th>
          </tr>
        </thead>
        <tbody>
          <tr class="active-row" v-for="item in listFail">
            <td>{{ item.style }}</td>
            <td>{{ item.color }}</td>
            <td>{{ item.size }}</td>
            <td>{{ item.quatity }}</td>
            <td>{{ item.pkg }}</td>
            <td>{{ item.dc }}</td>
            <td>{{ item.priority }}</td>
            <td>{{ item.revision }}</td>
            <td>{{ item.duedate.toISOString() }}</td>
          </tr>
        </tbody>
      </table>
    </div>


    <transition name="fade" mode="out-in" v-show="showSubmitFeedback && listFail.length <= 0">
      <div class="column">
        <button type="button" class="button1" v-show="submitable" v-on:click="submitdate">Fill Data</button>

        <b>Total: {{ list.length }} items</b>
        <table class="styled-table">
          <thead>
            <tr>
              <th class="text-left">
                Style
              </th>
              <th class="text-left">
                Color
              </th>
              <th class="text-left">
                Size
              </th>
              <th class="text-left">
                Quatity
              </th>
              <th class="text-left">
                PKG Style
              </th>
              <th class="text-left">
                DC
              </th>
              <th class="text-left">
                Priority
              </th>
              <th class="text-left">
                Revision
              </th>
              <th class="text-left">
                Sew Due
              </th>
            </tr>
          </thead>
          <tbody>
            <tr class="active-row" v-for="item in list">
              <td>{{ item.style }}</td>
              <td>{{ item.color }}</td>
              <td>{{ item.size }}</td>
              <td>{{ item.quatity }}</td>
              <td>{{ item.pkg }}</td>
              <td>{{ item.dc }}</td>
              <td>{{ item.priority }}</td>
              <td>{{ item.revision }}</td>
              <td>{{ item.duedate.toISOString() }}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </transition>
  </div>
  <loading v-model:active="isLoading" :can-cancel="false" :is-full-page="true" />
</template>

<style src="@vueform/multiselect/themes/default.css"></style>
<style scoped>
.custom-checkbox input[type="checkbox"] {
  appearance: none;
  -webkit-appearance: none;
  width: 20px;
  height: 20px;
  background-color: white;
  border-radius: 50%;
  border: 2px solid #fb0909;
  margin-bottom: -3px;
  margin-right: 5px;
}

.custom-checkbox input[type="checkbox"]:checked {
  background-color: #999;
}

.custom-checkbox label {
  color: #fb0909;
}

.margin-bottom {
  margin-bottom: 15px;
}

.fade-enter,
.fade-leave-active {
  opacity: 0;
}

.fade-enter-active,
.fade-leave-active {
  transition: opacity .5s;
}

* {
  box-sizing: border-box;
}

/* Style the search field */
form.example input[type=text] {
  padding: 10px;
  font-size: 17px;
  border: 1px solid grey;
  float: left;
  width: 100%;
  background: #f1f1f1;
}

/* Style the submit button */
form.example button {
  float: left;
  width: 20%;
  padding: 10px;
  width: 100%;
  margin-top: 10px;
  background: #2196F3;
  color: white;
  font-size: 17px;
  border: 1px solid grey;
  border-left: none;
  /* Prevent double borders */
  cursor: pointer;
}

form.example button:hover {
  background: #0b7dda;
}

/* Clear floats */
form.example::after {
  content: "";
  clear: both;
  display: table;
}

.button1 {
  display: inline-block;
  padding: 15px 25px;
  margin: 30px 0 30px 0;
  font-size: 24px;
  width: 100%;
  cursor: pointer;
  text-align: center;
  text-decoration: none;
  outline: none;
  color: #fff;
  background-color: #4CAF50;
  border: none;
  border-radius: 15px;
  box-shadow: 0 9px #999;
}

.button1:hover {
  background-color: #3e8e41
}

.button1:active {
  background-color: #3e8e41;
  box-shadow: 0 5px #666;
  transform: translateY(4px);
}



.styled-table {
  border-collapse: collapse;
  margin: 25px 0;
  font-size: 0.9em;
  font-family: sans-serif;
  min-width: 400px;
  box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}

.styled-table2 {
  border-collapse: collapse;
  margin: 25px 0;
  font-size: 0.9em;
  font-family: sans-serif;
  min-width: 400px;
  box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}





.styled-table2 thead tr {
  background-color: #981e00;
  color: #ffffff;
  text-align: left;
}

.styled-table2 th,
.styled-table2 td {
  padding: 12px 15px;
}


.styled-table2 tbody tr {
  border-bottom: 1px solid #dddddd;
}

.styled-table2 tbody tr:nth-of-type(even) {
  background-color: #f3f3f3;
}

.styled-table2 tbody tr:last-of-type {
  border-bottom: 2px solid #009879;
}

.styled-table2 tbody tr.active-row {
  font-weight: bold;
  color: #009879;
}


.styled-table2 thead tr {
  background-color: #b94f09;
  color: #ffffff;
  text-align: left;
}

.styled-table th,
.styled-table td {
  padding: 12px 15px;
}


.styled-table tbody tr {
  border-bottom: 1px solid #dddddd;
}

.styled-table tbody tr:nth-of-type(even) {
  background-color: #f3f3f3;
}

.styled-table tbody tr:last-of-type {
  border-bottom: 2px solid #009879;
}

.styled-table tbody tr.active-row {
  font-weight: bold;
  color: #009879;
}

* {
  box-sizing: border-box;
}

input[type=text],
select,
textarea {
  width: 100%;
  padding: 12px;
  border: 1px solid #ccc;
  border-radius: 4px;
  resize: vertical;
}

label {
  padding: 12px 12px 12px 0;
  display: inline-block;
}

input[type=submit] {
  background-color: #04AA6D;
  color: white;
  padding: 12px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  float: right;
}

input[type=submit]:hover {
  background-color: #45a049;
}

.container {
  border-radius: 5px;
  background-color: #f2f2f2;
  padding: 20px;
}

.col-25 {
  float: left;
  width: 25%;
  margin-top: 6px;
}

.col-75 {
  float: left;
  width: 75%;
  margin-top: 6px;
}

/* Clear floats after the columns */
.row::after {
  content: "";
  display: table;
  clear: both;
}

/* Responsive layout - when the screen is less than 600px wide, make the two columns stack on top of each other instead of next to each other */
@media screen and (max-width: 600px) {

  .col-25,
  .col-75,
  input[type=submit] {
    width: 100%;
    margin-top: 0;
  }
}

input[type=file]::file-selector-button {
  border: 2px solid #5c6ae7;
  padding: 12px;
  border-radius: 4px;
  background-color: #a29bfe;
  transition: 1s;
}

input[type=file]::file-selector-button:hover {
  background-color: #81ecec;
  border: 2px solid #00cec9;
}
</style>
