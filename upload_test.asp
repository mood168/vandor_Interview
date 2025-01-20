<%@LANGUAGE="VBSCRIPT" codepage="65001"%>
<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="Mark Otto, Jacob Thornton, and Bootstrap contributors">
    <meta name="generator" content="Hugo 0.79.0">
    <title>7-ELEVEN 詐騙申訴平臺</title>

    <link rel="canonical" href="https://bootstrap5.hexschool.com/docs/5.0/examples/checkout/">

  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>   

    <!-- Bootstrap core CSS -->
<link href="css/bootstrap.min.css" rel="stylesheet">
    <!-- Favicons -->
<link rel="apple-touch-icon" href="/docs/5.0/assets/img/favicons/apple-touch-icon.png" sizes="180x180">
<link rel="icon" href="/docs/5.0/assets/img/favicons/favicon-32x32.png" sizes="32x32" type="image/png">
<link rel="icon" href="/docs/5.0/assets/img/favicons/favicon-16x16.png" sizes="16x16" type="image/png">
<link rel="manifest" href="/docs/5.0/assets/img/favicons/manifest.json">
<link rel="mask-icon" href="/docs/5.0/assets/img/favicons/safari-pinned-tab.svg" color="#7952b3">
<link rel="icon" href="/docs/5.0/assets/img/favicons/favicon.ico">
<meta name="theme-color" content="#7952b3">
<style>
body {font-family: Arial, Helvetica, sans-serif;}

/* The Modal (background) */
.modal {
  display: none; /* Hidden by default */
  position: fixed; /* Stay in place */
  z-index: 99; /* Sit on top */
  padding-top: 100px; /* Location of the box */
  left: 0;
  top: 0;
  width: 100%; /* Full width */
  height: 100%; /* Full height */  
  overflow: auto; /* Enable scroll if needed */
  background-color: rgb(0,0,0); /* Fallback color */
  background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
}

/* Modal Content */
.modal-content {
  background-color: #fefefe;
  margin: auto;
  padding: 20px;
  border: 1px solid #888;
  width: 50%;
}

/* The Close Button */
.close {
  color: #aaaaaa;
  float: right;
  font-size: 18px;
  font-weight: bold;
}

.close:hover,
.close:focus {
  color: #000;
  text-decoration: none;
  cursor: pointer;
}

.closealt {
  color: #aaaaaa;
  float: right;
  font-size: 18px;
  font-weight: bold;
}

.closealt:hover,
.closealt:focus {
  color: #000;
  text-decoration: none;
  cursor: pointer;
}
</style>
<style>
  .bd-placeholder-img {
	font-size: 1.125rem;
	text-anchor: middle;
	-webkit-user-select: none;
	-moz-user-select: none;
	user-select: none;
  }

  @media (min-width: 768px) {
	.bd-placeholder-img-lg {
	  font-size: 3.5rem;
	}
  }
</style>
<!-- Custom styles for this template -->
<link href="form-validation.css" rel="stylesheet">
<link href="https://bootstrap5.hexschool.com/docs/5.1/examples/modals/modals.css" rel="stylesheet">
<style>
       
        .process-container {
            display: flex;
            justify-content: space-between;
            height: 80px;
            gap: 10px;
        }
        .process-step {
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .icon {
            width: 30px;
            height: 30px;
            background-size: contain;
            background-repeat: no-repeat;
            background-position: center;
            margin-bottom: 10px; /* 圖標與box的距離 */
        }
        .box {
		  display: flex;
		  align-items: center;
		  border: 2px solid #878787;
		  border-radius: 4px;
		  padding: 5px;
		  width: 100%;
		  height: 50px;
		}

		.icon-container {
		  flex: 0 0 auto;
		  width: 50px; /* 調整圖標容器的寬度 */
		  margin-right: 5px; /* 圖標和文字之間的間距 */
		  align-items: center;
		}

		.icon {
		  width: 100%;
		  height: auto;
		  max-height: 50; /* 確保圖標不會太大 */
		}

		.text-container {
		  flex: 1;
		  font-size: 12px;
		  font-weight: bold;
		  display: flex;
		  align-items: center;
		}
    </style>
   
</head>
<body class="bg-light">
    <form action="bankinfoupload.asp" name="text1" method="post" enctype="multipart/form-data" accept-charset="UTF-8" class="row g-3 needs-validation" novalidate>
        <div class="row g-3 align-items-end">
            <div class="col-sm-6">
            <label for="bank" id="bk" class="form-label" style="display: flex;background-color: #e8ecef;">銀行代碼&nbsp;<span class="badge bg-secondary rounded-pill">*必選</span></label>
                    <select id="bank" class="form-select" name="bank" onchange="detectChange(this)" required>
                    <option value="">請選擇銀行</option>
                    <option value="004">臺灣銀行-004</option>
                    <option value="005">臺灣土地銀行-005</option>
                    <option value="006">合作金庫商業銀行-006</option>
                    <option value="007">第一商業銀行-007</option>
                    <option value="008">華南商業銀行-008</option>
                    <option value="009">彰化商業銀行-009</option>
                    <option value="011">上海商業儲蓄銀行-011</option>
                    <option value="012">台北富邦商業銀行-012</option>
                    <option value="013">國泰世華商業銀行-013</option>
                    <option value="015">中國輸出入銀行-015</option>
                    <option value="016">高雄銀行-016</option>
                    <option value="017">兆豐國際商業銀行-017</option>
                    <option value="021">花旗(台灣)商業銀行-021</option>
                    <option value="048">王道商業銀行-048</option>
                    <option value="050">臺灣中小企業銀行-050</option>
                    <option value="052">渣打國際商業銀行-052</option>
                    <option value="053">台中商業銀行-053</option>
                    <option value="054">京城商業銀行-054</option>
                    <option value="081">滙豐(台灣)商業銀行-081</option>
                    <option value="101">瑞興商業銀行-101</option>
                    <option value="102">華泰商業銀行-102</option>
                    <option value="103">臺灣新光商業銀行-103</option>
                    <option value="108">陽信商業銀行-108</option>
                    <option value="118">板信商業銀行-118</option>
                    <option value="147">三信商業銀行-147</option>
                    <option value="700">中華郵政 郵局-700</option>
                    <option value="803">聯邦商業銀行-803</option>
                    <option value="805">遠東國際商業銀行-805</option>
                    <option value="806">元大商業銀行-806</option>
                    <option value="807">永豐商業銀行-807</option>
                    <option value="808">玉山商業銀行-808</option>
                    <option value="809">凱基商業銀行-809</option>
                    <option value="810">星展(台灣)商業銀行-810</option>
                    <option value="812">台新國際商業銀行-812</option>
                    <option value="815">日盛國際商業銀行-815</option>
                    <option value="816">安泰商業銀行-816</option>
                    <option value="822">中國信託商業銀行-822</option>
                    <option value="823">將來商業銀行(Next Bank)-823</option>
                    <option value="824">連線商業銀行(LINE Bank)-824</option>
                    <option value="826">樂天國際商業銀行-826</option>				
                </select>
            </div>
            <div class="col-sm-2">
            <label for="branchname" id="bn" class="form-label" style="display: flex;background-color: #e8ecef;">分行名稱&nbsp;<span class="badge bg-secondary rounded-pill">*必填</span></label>
            <input type="text" name="branchname" class="form-control input-group" id="branchname" pattern="^[\u4e00-\u9fa5]+$" size="5" maxlength="12" placeholder="中文" value="" required>
            <div class="invalid-feedback">
                *必填中文
            </div>
            </div>
            <div class="col-sm-4">
            <label for="bankaccount" id="ba" class="form-label" style="display: flex;background-color: #e8ecef;">銀行帳號&nbsp;<span class="badge bg-secondary rounded-pill">*必填</span></label>
            <input type="text" name="bankaccount" class="form-control input-group" id="bankaccount" pattern="^([0-9]{12}|[0-9]{13}|[0-9]{14})$" maxlength="14" size="5" placeholder="12-14碼數字" value="" required>
            <div class="invalid-feedback">
                *必填12-14碼數字
            </div>
            </div>
            <div class="col-sm-5">
            <label for="accountname" id="ac" class="form-label" style="display: flex;background-color: #e8ecef;">戶名&nbsp;<span class="badge bg-secondary rounded-pill">*必填</span></label>
            <input type="text" name="accountname" class="form-control input-group" id="accountname" pattern="^[\u4e00-\u9fa5]+$|^[a-zA-Z\s]+$" size="5" maxlength="12" placeholder="中文或英文" value="" required>
            <div class="invalid-feedback">
                *必填中文或英文
            </div>
            </div>
            <div class="col-sm-5">
            <label for="cfile" id="cf" class="form-label" style="display: flex;background-color: #e8ecef;">存摺正面拍照上傳&nbsp;<span class="badge bg-secondary rounded-pill">*必選</span></label>
            <input type="file" name="file1" class="form-control input-group" id="cfile" size="5" placeholder="選擇上傳圖檔 .jpg, .png" value="" required>
            <div class="invalid-feedback">
                *必選 選擇上傳圖檔 .jpg, .png
            </div>
            </div>
            <div class="col-sm-2">
            <input type="hidden" name="index_no" value="12345678901234">
            <button type="submit" class="btn btn-secondary btn-md">上傳並送出</button>
            </div>
        </div>						
    </form>
</body></html>