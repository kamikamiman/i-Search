// ------------------------------------------------------ //
//  PDFを結合する                                           //
// ------------------------------------------------------ //

async function PdfMerge() {  
    
  const folder     = DriveApp.getFolderById(mergeman); // フォルダを取得する
  const pdflist    = folder.getFiles();                // フォルダ内のPDFをリスト化
  const fileRename = `${rename}.pdf`;                  // 結合後ファイル名の指定
  let mergearr     = [];                               // リスト用の配列
  
  //mergearrにPDFデータをpush
  while (pdflist.hasNext()){
    
    const file   = pdflist.next();                     // ファイルを取得
    const fileName = file.getName();                   // ファイル名を取得
    const fileNum  = fileName.indexOf(rename);         // 変数renameの文字列がファイル名に存在するか
    const PDF = file.getMimeType() === MimeType.PDF;   // ファイルタイプがpdf
    
    // アップロードファイル名を含むpdfファイルを配列に格納
    if ( fileNum != -1 && PDF ) mergearr.push(file);
    
  };
  
  const mergedFile = mergePdfs(folder, fileRename, mergearr);  //PDF結合実行
  
  await FileMove(); // 結合したpdfを指定フォルダに移動

};


// ここからは完全コピー
function mergePdfs(directory, name, pdf_list) {
  
  if (name.slice(-4) != '.pdf') {
    
    name = name + '.pdf';
    
  }
  var newObjects = ['1 0 obj\r\n<</Type/Catalog/Pages 2 0 R >>\r\nendobj'];
  var pageAddresses = [];
  for (var argumentIndex in pdf_list) {
    
    var bytes = pdf_list[argumentIndex].getBlob().getBytes();
    var xrefByteOffset = '';
    var byteIndex = bytes.length - 1;
    while (!/\sstartxref\s/.test(xrefByteOffset)) {
      
      xrefByteOffset = String.fromCharCode(bytes[byteIndex]) + xrefByteOffset;
      byteIndex--;
      
    }
    xrefByteOffset = +(/\s\d+\s/.exec(xrefByteOffset)[0]);
    var objectByteOffsets = [];
    var trailerDictionary = '';
    var rootAddress = '';
    do {
      
      var xrefTable = '';
      var trailerEndByteOffset = byteIndex;
      byteIndex = xrefByteOffset;
      for (byteIndex; byteIndex <= trailerEndByteOffset; byteIndex++) {
        
        xrefTable = xrefTable + String.fromCharCode(bytes[byteIndex]);
      }
      xrefTable = xrefTable.split(/\s*trailer\s*/);
      trailerDictionary = xrefTable[1];
      if (objectByteOffsets.length < 1) {
        
        rootAddress = /\d+\s+\d+\s+R/.exec(/\/Root\s*\d+\s+\d+\s+R/.exec(trailerDictionary)[0])[0].replace('R', 'obj');
      }
      xrefTable = xrefTable[0].split('\n');
      xrefTable.shift();
      while (xrefTable.length > 0) {
        
        var xrefSectionHeader = xrefTable.shift().split(/\s+/);
        var objectNumber = +xrefSectionHeader[0];
        var numberObjects = +xrefSectionHeader[1];
        for (var entryIndex = 0; entryIndex < numberObjects; entryIndex++) {
          
          var entry = xrefTable.shift().split(/\s+/);
          objectByteOffsets.push([[objectNumber, +entry[1], 'obj'], +entry[0]]);
          objectNumber++;
          
        }
        
      }
      if (/\s*\/Prev/.test(trailerDictionary)) {
        
        xrefByteOffset = +(/\s*\d+\s/.exec(/\s*\/Prev\s*\d+\s/.exec(trailerDictionary)[0])[0]);
        
      }
      
    } while (/\s*\/Prev/.test(trailerDictionary));
    var rootObject = getObject(rootAddress, objectByteOffsets, bytes);
    var pagesAddress = /\d+\s+\d+\s+R/.exec(/\/Pages\s*\d+\s+\d+\s+R/.exec(rootObject)[0])[0].replace('R', 'obj');
    var pagesObject = getObject(pagesAddress, objectByteOffsets, bytes);
    var objects = getDependencies(pagesObject, objectByteOffsets, bytes);
    var newObjectsInsertionIndex = newObjects.length;
    for (var objectIndex = 0; objectIndex < objects.length; objectIndex++) {
      
      var newObjectAddress = [(newObjects.length + 3) + '', 0 + '', 'obj'];
      if (!Array.isArray(objects[objectIndex])) {
        
        objects[objectIndex] = [objects[objectIndex]];
          
      }
      objects[objectIndex].unshift(newObjectAddress);
      var objectAddress = objects[objectIndex][1].match(/\d+\s+\d+\s+obj/)[0].split(/\s+/);
      objects[objectIndex].splice(1, 0, objectAddress);
      if (/\/Type\s*\/Page[^s]/.test(objects[objectIndex][2])) {
        
        objects[objectIndex][2] = objects[objectIndex][2].replace(/\/Parent\s*\d+\s+\d+\s+R/.exec(objects[objectIndex][2])[0], '/Parent 2 0 R');
        pageAddresses.push(newObjectAddress.join(' ').replace('obj', 'R'));
        
      }
      var addressRegExp = new RegExp(objectAddress[0] + '\\s+' + objectAddress[1] + '\\s+' + 'obj');
      objects[objectIndex][2] = objects[objectIndex][2].replace(addressRegExp.exec(objects[objectIndex][2])[0], newObjectAddress.join(' '));
      newObjects.push(objects[objectIndex]);
      
    }
    for (var referencingObjectIndex = newObjectsInsertionIndex; referencingObjectIndex < newObjects.length; referencingObjectIndex++) {
      
      var references = newObjects[referencingObjectIndex][2].match(/\d+\s+\d+\s+R/g);
      if (references != null) {
        
        var string = newObjects[referencingObjectIndex][2];
        var referenceIndices = [];
        var currentIndex = 0;
        for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {
          
          referenceIndices.push([]);
          referenceIndices[referenceIndex].push(string.slice(currentIndex).indexOf(references[referenceIndex]) + currentIndex);
          referenceIndices[referenceIndex].push(references[referenceIndex].length);
          currentIndex += string.slice(currentIndex).indexOf(references[referenceIndex]);
          
        }
        for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {
          
          var objectAddress = references[referenceIndex].replace('R', 'obj').split(/\s+/);
          for (var objectIndex = newObjectsInsertionIndex; objectIndex < newObjects.length; objectIndex++) {
            
            if (arrayEquals(objectAddress, newObjects[objectIndex][1])) {
              
              var length = string.length;
              newObjects[referencingObjectIndex][2] = string.slice(0, referenceIndices[referenceIndex][0]) + newObjects[objectIndex][0].join(' ').replace('obj', 'R') +
                string.slice(referenceIndices[referenceIndex][0] + referenceIndices[referenceIndex][1]);
              string = newObjects[referencingObjectIndex][2];
              var newLength = string.length;
              if (!(length == newLength)) {
                
                for (var subsequentReferenceIndex = referenceIndex + 1; subsequentReferenceIndex < references.length; subsequentReferenceIndex++) {
                  
                  referenceIndices[subsequentReferenceIndex][0] += (newLength - length);
                  
                }
                
              }
              break;
              
            }
            
          }
          
        }
        
      }
      
    }
    for (var objectIndex = newObjectsInsertionIndex; objectIndex < newObjects.length; objectIndex++) {
      
      if (Array.isArray(newObjects[objectIndex])) {
        
        if (newObjects[objectIndex][3] != undefined) {
          
          newObjects[objectIndex] = newObjects[objectIndex].slice(2);
          
        } else {
          
          newObjects[objectIndex] = newObjects[objectIndex][2];
          
        }
        
      }
      
    }
    
  }
  newObjects.splice(1, 0, '2 0 obj\r\n<</Type/Pages/Count ' + pageAddresses.length + ' /Kids [' + pageAddresses.join(' ') + ' ]>>\r\nendobj');
  newObjects.splice(2, 0, '3 0 obj\r\n<</Title (' + name + ') /Producer (PdfManipulation.mergePdfs\\(\\), a Google Apps Script project by Jarom Luker \\(pricebook@hbboys.com\\)) /CreationDate (D' +
       Utilities.formatDate(new Date(), CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00) /ModDate (D" + Utilities.formatDate(new Date(),
       CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00)>>\r\nendobj");
  var byteOffsets = [0];
  var bytes = [];
  var header = '%PDF-1.3\r\n';
  for (var headerIndex = 0; headerIndex < header.length; headerIndex++) {
    
    bytes.push(header.charCodeAt(headerIndex));
    
  }
  bytes.push('%'.charCodeAt(0));
  for (var characterCode = -127; characterCode < -123; characterCode++) {
    
    bytes.push(characterCode);
    
  }
  bytes.push('\r'.charCodeAt(0));
  bytes.push('\n'.charCodeAt(0));
  while (newObjects.length > 0) {
    
    byteOffsets.push(bytes.length);
    var object = newObjects.shift();
    if (Array.isArray(object)) {
      
      var streamKeyword = /stream\s*\n/.exec(object[0])[0];
      if (streamKeyword.indexOf('\n\n') > streamKeyword.length - 3) {
        
        streamKeyword = streamKeyword.slice(0, -1);
        
      } else if (streamKeyword.indexOf('\r\n\r\n') > streamKeyword.length - 5) {
        
        streamKeyword = streamKeyword.slice(0, -2);
        
      }
      var streamIndex = object[0].indexOf(streamKeyword) + streamKeyword.length;
      for (var objectIndex = 0; objectIndex < streamIndex; objectIndex++) {
        
        bytes.push(object[0].charCodeAt(objectIndex))
        
      }
      bytes = bytes.concat(object[1]);
      for (var objectIndex = streamIndex; objectIndex < object[0].length; objectIndex++) {
        
        bytes.push(object[0].charCodeAt(objectIndex));
        
      }
      
    } else {
      
      for (var objectIndex = 0; objectIndex < object.length; objectIndex++) {
        
        bytes.push(object.charCodeAt(objectIndex));
        
      }
      
    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));
    
  }
  var xrefByteOffset = bytes.length;
  var xrefHeader = 'xref\r\n';
  for (var xrefHeaderIndex = 0; xrefHeaderIndex < xrefHeader.length; xrefHeaderIndex++) {
    
    bytes.push(xrefHeader.charCodeAt(xrefHeaderIndex));
    
  }
  var xrefSectionHeader = '0 ' + byteOffsets.length + '\r\n';
  for (var xrefSectionHeaderIndex = 0; xrefSectionHeaderIndex < xrefSectionHeader.length; xrefSectionHeaderIndex++) {
    
    bytes.push(xrefSectionHeader.charCodeAt(xrefSectionHeaderIndex));
    
  }
  for (var byteOffsetIndex = 0; byteOffsetIndex < byteOffsets.length; byteOffsetIndex++) {
    
    for (var byteOffsetStringIndex = 0; byteOffsetStringIndex < 10; byteOffsetStringIndex++) {
      
      bytes.push(Utilities.formatString('%010d', byteOffsets[byteOffsetIndex]).charCodeAt(byteOffsetStringIndex));
      
    }
    bytes.push(' '.charCodeAt(0));
    if (byteOffsetIndex == 0) {
      
      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {
        
        bytes.push('65535'.charCodeAt(generationStringIndex));
        
      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {
        
        bytes.push(' f'.charCodeAt(keywordIndex));
        
      }
      
    } else {
      
      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {
        
        bytes.push('0'.charCodeAt(0));
        
      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {
        
        bytes.push(' n'.charCodeAt(keywordIndex));
        
      }
    
    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));
    
  }
  for (var trailerHeaderIndex = 0; trailerHeaderIndex < 9; trailerHeaderIndex++) {
    
    bytes.push('trailer\r\n'.charCodeAt(trailerHeaderIndex));
    
  }
  var idBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (new Date).toString());
  var id = '';
  for (var idByteIndex = 0; idByteIndex < idBytes.length; idByteIndex++) {
    
    id = id + ('0' + (idBytes[idByteIndex] & 0xFF).toString(16)).slice(-2);
    
  }
  var trailer = '<</Size ' + (byteOffsets.length) + ' /Root 1 0 R /Info 2 0 R /ID [<' + id + '> <' + id + '>]>>\r\nstartxref\r\n' + xrefByteOffset + '\r\n%%EOF';
  for (var trailerIndex = 0; trailerIndex < trailer.length; trailerIndex++) {
    
    bytes.push(trailer.charCodeAt(trailerIndex));
    
  }
  return directory.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  function getObject(objectAddress, objectByteOffsets, bytes) {
    
    objectAddress = objectAddress.split(/\s+/);
    for (var addressIndex = 0; addressIndex < 2; addressIndex++) {
      
      objectAddress[addressIndex] = +objectAddress[addressIndex];
      
    }
    var object = [];
    var byteIndex = 0;
//    for each (var offset in objectByteOffsets) {
      for (var key in objectByteOffsets) {
        offset = objectByteOffsets[key]

      
      if (arrayEquals(objectAddress, offset[0])) {
        
        byteIndex = offset[1];
        break;
        
      }
      
    }
    object.push('');
    while (object[0].indexOf('endobj') <= -1) {
      
      if (/stream\s*\n/.test(object[0])) {
        
        var streamLength;
        var lengthFinder = object[0].slice(object[0].indexOf(/\/Length/.exec(object[0])[0]));
        if (/\/Length\s*\d+\s+\d+\s+R/.test(lengthFinder)) {
          
          var lengthObjectAddress = /\d+\s+\d+\s+R/.exec(/\/Length\s*\d+\s+\d+\s+R/.exec(lengthFinder)[0])[0].split(/\s+/);
          lengthObjectAddress[2] = 'obj';
          for (var addressIndex = 0; addressIndex < 2; addressIndex++) {
            
            lengthObjectAddress[addressIndex] = +lengthObjectAddress[addressIndex];
            
          }
          var lengthObject = ''
          var lengthByteIndex = 0;
//          for each (var offset in objectByteOffsets) {
            for (var key in objectByteOffsets) {
              offset = objectByteOffsets[key]
            
            if (arrayEquals(lengthObjectAddress, offset[0])) {
              
              lengthByteIndex = offset[1];
              break;
              
            }
            
          }
          while (lengthObject.indexOf('endobj') <= -1) {
            
            lengthObject = lengthObject + String.fromCharCode(bytes[lengthByteIndex]);
            lengthByteIndex++;
            
          }
          streamLength = +(lengthObject.match(/obj\s*\n\s*\d+\s*\n\s*endobj/)[0].match(/\d+/)[0]);
          
        } else {
          
          streamLength = +(/\d+/.exec(lengthFinder)[0]);
          
        }
        var streamBytes = bytes.slice(byteIndex, byteIndex + streamLength);
        object.push(streamBytes);
        byteIndex += streamLength;
        while (object[0].indexOf('endobj') <= -1) {
          
          object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
          byteIndex++;
          
        }
        return object;
        
      }
      object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
      byteIndex++;
      
    }
    return object[0];
    
  }
  function arrayEquals(array1, array2) {
    
    if (array1 == array2) {
      
      return true;
      
    }
    if (array1 == null && array2 == null) {
      
      return true;
      
    } else if (array1 == null || array2 == null) {
      
      return false;
      
    }
    if (array1.length != array2.length) {
      
      return false;
      
    }
    for (var index = 0; index < array1.length; index++) {
      
      if (Array.isArray(array1[index])) {
        
        if (!arrayEquals(array1[index], array2[index])) {
          
          return false;
          
        }
        continue;
        
      }
      if (array1[index] != array2[index]) {
        
        return false;
        
      }
      
    }
    return true;
    
  }
  function getDependencies(objectString, objectByteOffsets, bytes) {
    
    var dependencies = [];
    var references = objectString.match(/\d+\s+\d+\s+R/g);
    if (references != null) {
      
      while (references.length > 0) {
        
        if (/\/Parent/.test(objectString.slice(objectString.indexOf(references[0]) - 8, objectString.indexOf(references[0])))) {
          
          references.shift();
          continue;
          
        }
        var dependency = getObject(references.shift().replace('R', 'obj'), objectByteOffsets, bytes);
        var dependencyExists = false;
//        for each (var entry in dependencies) {
          for (var key in dependencies) {
            entry = dependencies[key]

          
          dependencyExists = (arrayEquals(dependency, entry)) ? true : dependencyExists;
          
        }
        if (!dependencyExists) {
          
          dependencies.push(dependency);
          
        }
        if (Array.isArray(dependency)) {
          
          dependencies = dependencies.concat(getDependencies(dependency[0], objectByteOffsets, bytes));
          
        } else {
          
          dependencies = dependencies.concat(getDependencies(dependency, objectByteOffsets, bytes));
          
        }
        
      }
      
    }
    return dependencies;
    
  } 
}
