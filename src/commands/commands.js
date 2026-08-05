
var flag = 0;
var auth_obj;
var tname = "";
Office.onReady((info) => {
  // If needed, Office.js is ready to be called
  if (info.host == Office.HostType.Excel) {
  }
});
function getParameterByName(name, url = window.location.href) {
  name = name.replace(/[\[\]]/g, "\\$&");
  var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
    results = regex.exec(url);
  if (!results) return null;
  if (!results[2]) return "";
  return decodeURIComponent(results[2].replace(/\+/g, " "));
}
window.addEventListener("load", async () => {
  document.getElementById("fileInput").addEventListener("change", function () {
    const file = this.files[0];
    if (file.name.split(".").pop() == "ionapi") {
      if (!file) {
        return;
      }
    }
    const reader = new FileReader();
    reader.onload = function (e) {
      const content = e.target.result;
      auth_obj = JSON.parse(e.target.result);
      tname = file.name;
      const data = {
        name: file.name,
        size: file.size,
        content: content,
      };

      // Load existing files or create empty list
      let fileList = JSON.parse(localStorage.getItem("uploadedFiles")) || [];

      // Avoid duplicates
      const exists = fileList.find((f) => f.name === data.name);
      if (!exists) {
        fileList.push(data);
        localStorage.setItem("uploadedFiles", JSON.stringify(fileList));
        updateFileList();
      }

      flag = 0;
      // Using For Each Loop
      Object.keys(localStorage).every(function (key) {
        if (key.includes("ionAPI")) {
          let opt = JSON.parse(localStorage.getItem(key));
          if (opt.cn == auth_obj.cn) {
            // localStorage.removeItem(key);
            localStorage.setItem(key, JSON.stringify(auth_obj));
            flag = 1;

            // Display Modal
            // var myModal = new bootstrap.Modal(document.getElementById("myModal"));
            // document.getElementById("modalHeading").innerHTML = "Upload Profile";
            // document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${auth_obj.ti} tenant.`;
            // myModal.show();

            return false;
          }
        }
        return true;
      });
      // If User Uploads a New File
      if (flag == 0) {
        flag = 1;
        var date = new Date();
        var timestamp =
          date.getDate().toString() +
          (date.getMonth() + 1).toString() +
          date.getFullYear().toString() +
          date.getHours().toString() +
          date.getMinutes().toString() +
          date.getSeconds().toString();
        localStorage.setItem(`ionAPI_${timestamp}`, JSON.stringify(auth_obj));

        // Load List value
        var opt = localStorage.getItem(`ionAPI_${timestamp}`);
        var ul = document.getElementById("dynamic-list");
        var li = document.createElement("li");
        li.setAttribute("id", JSON.parse(opt).cn);
        ul.appendChild(li);

        var btn = document.createElement("button");
        btn.type = "button";
        btn.title = "Click to Sign In";
        btn.classList.add("btn");
        btn.classList.add("btn-link");
        btn.classList.add("btn-sm");
        btn.innerHTML = JSON.parse(opt).cn;
        li.appendChild(btn);
        li.innerHTML += `<span class="close"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16"
  fill="currentColor" class="bi bi-x-circle-fill" viewBox="0 0 16 16">
  <path
      d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM5.354 4.646a.5.5 0 1 0-.708.708L7.293 8l-2.647 2.646a.5.5 0 0 0 .708.708L8 8.707l2.646 2.647a.5.5 0 0 0 .708-.708L8.707 8l2.647-2.646a.5.5 0 0 0-.708-.708L8 7.293 5.354 4.646z" />
</svg></span>
</li>`;
        setButtonEventListener();
        setCloseEventListener();

        // var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        // document.getElementById("modalHeading").innerHTML = "Upload Profile";
        // document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${JSON.parse(opt).ti} tenant.`;
        // myModal.show();
      }

      // Send selected file to parent (optional)
      //  Office.context.ui.messageParent(JSON.stringify(data));
    };

    if (file) {
      reader.readAsText(file);
    }
  });
  var codes = getParameterByName("code");
  if (codes != null) {
    console.log(codes);
    var pongurl = localStorage.getItem("iu") + "/" + localStorage.getItem("ti");
    // Send tenant details to taskpane
    var messageObject = {
      messageType: "tenant",
      tenant_name: localStorage.getItem("ti"),
      endpoint_url: localStorage.getItem("iu"),
      logout_url: localStorage.getItem("pu"),
    };
    var jsonMessage = JSON.stringify(messageObject);
    var urlencoded = new URLSearchParams();
    urlencoded.append("grant_type", "authorization_code");
    urlencoded.append("ci", localStorage.getItem("ci"));
    urlencoded.append("cs", localStorage.getItem("cs"));
    urlencoded.append("pu", localStorage.getItem("pu1"));
    urlencoded.append("ot", localStorage.getItem("ot"));
    urlencoded.append("code", codes);
    urlencoded.append("ru", localStorage.getItem("ru"));
    var authids =
      "ci==" +
      localStorage.getItem("ci") +
      "|cs==" +
      localStorage.getItem("cs") +
      "|ru==" +
      localStorage.getItem("ru") +
      "|authcode==" +
      codes +
      "|pu==" +
      localStorage.getItem("pu1") +
      "|ot==" +
      localStorage.getItem("ot") +
      "|apiurl==" +
      pongurl + "|ti==" + localStorage.getItem('ti');
    //var authids="ci=="+localStorage.getItem('ci')+"|cs=="+localStorage.getItem('cs')+"|ru=="+localStorage.getItem('ru')+"|authcode=="+codes+"|pu=="+localStorage.getItem('pu1')+"|ot=="+localStorage.getItem('ot');
    Office.context.ui.messageParent(authids);

    //Remove ci, cs and ti from localStorage
    localStorage.removeItem("ci");
    localStorage.removeItem("cs");
    localStorage.removeItem("ti");
    localStorage.removeItem("iu");
    localStorage.removeItem("pu");
    localStorage.removeItem("ot");
    localStorage.removeItem("ru");
    localStorage.removeItem("pu1");
  } else {
    // Remove CI, CS and TI if exits
    if ("ci" in localStorage) localStorage.removeItem("ci");
    if ("cs" in localStorage) localStorage.removeItem("cs");
    if ("ti" in localStorage) localStorage.removeItem("ti");
    if ("iu" in localStorage) localStorage.removeItem("iu");
    if ("pu" in localStorage) localStorage.removeItem("pu");
    if ("ot" in localStorage) localStorage.removeItem("ot");
    if ("ru" in localStorage) localStorage.removeItem("ru");
    if ("pu1" in localStorage) localStorage.removeItem("pu1");
  }
});
// Populate <select> with uploaded file names
function updateFileList() {
  // localStorage.clear();
  const ul = document.getElementById("dynamic-list");
  ul.innerHTML = "";
  const fileList = JSON.parse(localStorage.getItem("uploadedFiles")) || [];
  Object.keys(localStorage)
    .sort()
    .reverse()
    .forEach(function (key) {
      if (key.includes("ionAPI")) {
        // const files=fileList[0];
        const opt = localStorage.getItem(key);
        var name = JSON.parse(opt).cn;

        const li = document.createElement("li");
        li.setAttribute("id", name);
        li.innerHTML = `<button type="button" title="Click to Sign In" class="btn btn-link btn-sm">${name}</button>
    <span class="close"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16"
      fill="currentColor" class="bi bi-x-circle-fill" viewBox="0 0 16 16">
      <path
          d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM5.354 4.646a.5.5 0 1 0-.708.708L7.293 8l-2.647 2.646a.5.5 0 0 0 .708.708L8 8.707l2.646 2.647a.5.5 0 0 0 .708-.708L8.707 8l2.647-2.646a.5.5 0 0 0-.708-.708L8 7.293 5.354 4.646z" />
  </svg></span>
  </li>`;
        ul.appendChild(li);
      }
    });
  setButtonEventListener();
  setCloseEventListener();
}

function setCloseEventListener() {
  /* Get all elements with class="close" */
  closebtns = document.getElementsByClassName("close");
  var i;

  /* Loop through the elements, and hide the parent, when clicked on */
  for (i = 0; i < closebtns.length; i++) {
    closebtns[i].addEventListener("click", function () {
      this.parentElement.style.display = "none";
      const name = this.parentElement.id;
      const ul = document.getElementById("dynamic-list");
      if (this.parentElement && ul.contains(this.parentElement)) {
        ul.removeChild(this.parentElement);
      }
      // //ul.clear();
      // ul.innerHTML="";
      // // while (ul.firstChild) {
      // //     ul.removeChild(ul.firstChild); // ✅ this works because firstChild is always a Node
      // //   }
      // var item = document.getElementById(name);
      // ul.removeChild(item);
      Object.keys(localStorage).every(function (key) {
        if (key.includes("ionAPI")) {
          var opt = JSON.parse(localStorage.getItem(key));

          if (opt.cn == name) {
            localStorage.removeItem(key);
            return false;
          }
        }
        return true;
      });
    });
  }
}
function setButtonEventListener() {
  // Add listener to li elements
  items = document.getElementsByClassName("btn-link");
  var i;

  /* Loop through the elements, and hide the parent, when clicked on */
  for (i = 0; i < items.length; i++) {
    items[i].addEventListener("click", function () {
      logIn(this.parentElement.id);
    });
  }
}
function logIn(name) {
  Object.keys(localStorage).every(function (key) {
    if (key.includes("ionAPI")) {
      var auth_obj = JSON.parse(localStorage.getItem(key));

      if (auth_obj.cn == name) {
        // Set CI , CS and TI
        localStorage.setItem(`ci`, auth_obj.ci);
        localStorage.setItem(`cs`, auth_obj.cs);
        localStorage.setItem(`ti`, auth_obj.ti);
        localStorage.setItem(`iu`, auth_obj.iu);
        localStorage.setItem(`ot`, auth_obj.ot);
        localStorage.setItem(`ru`, auth_obj.ru);
        localStorage.setItem(`pu1`, auth_obj.pu);

        // Set Logout URL from pu
        var url = auth_obj.pu.search(".com");
        url = auth_obj.pu.slice(0, url + 4);
        localStorage.setItem(`pu`, url);
        window.location.replace(
          `${auth_obj["pu"]}${auth_obj["oa"]}?client_id=${auth_obj["ci"]}&AuthMode=Prompt&response_type=code&redirect_uri=${auth_obj["ru"]}&TenantId=${auth_obj["ti"]}`
        );

        return false;
      }
    }
    return true;
  });
}
// On page load, populate list if data exists
document.addEventListener("click", function (e) {
  if (e.target.classList.contains("btn-link")) {
    logIn(e.target.parentElement?.id);
  }
  if (e.target.classList.contains("close")) {
    var close = e.target.parentElement?.id;
    close.style.display = "none";
    var name = close;
    var ul = document.getElementById("dynamic-list");
    // ul.clear();
    ul.innerHTML = "";
    var item = document.getElementById(name);
    ul.removeChild(item);
    Object.keys(localStorage).every(function (key) {
      if (key.includes("ionAPI")) {
        var opt = JSON.parse(localStorage.getItem(key));

        if (opt.cn == name) {
          localStorage.removeItem(key);
          return false;
        }
      }
      return true;
    });
  }
});
window.onload = updateFileList;
