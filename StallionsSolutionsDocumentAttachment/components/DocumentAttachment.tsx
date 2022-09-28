import React, { useEffect, useState } from 'react';
import { Col, Container, Row } from 'react-bootstrap';
import { FileUploader } from 'react-drag-drop-files';
import { saveAs } from 'file-saver';

import 'bootstrap/dist/css/bootstrap.min.css';
import "./style.css";

const DocumentAttachment = (props: any) => {
    const [attachments, setAttachments] = useState<any>([]);
    const [file, setFile] = useState("");
    const [base64Str, setBase64] = useState("");
    const entityId = props.componentContext.mode.contextInfo.entityId;
    const logicalName: string = props.componentContext.mode.contextInfo.entityTypeName;

    useEffect(() => {
        props.componentContext.factory.requestRender();

        const searchQuery = "?$select=annotationid,documentbody,mimetype,notetext,subject,filename,filesize,createdon" +
            "&$filter=_objectid_value eq " +
            entityId +
            " and isdocument eq true " +
            "&$orderby=createdon asc";
        props.componentContext.webAPI.retrieveMultipleRecords("annotation", searchQuery)
            .then((res: any) => {
                setAttachments(res.entities);

                if (res.entities.length > 0 && res.entities[0].mimetype === "application/pdf") {
                    const documentbody = res.entities[0].documentbody;
                    const contentType = res.entities[0].mimetype;
                    const blob = b64toBlob(documentbody, contentType, 512);
                    const objectURL = URL.createObjectURL(blob);

                    setFile(objectURL);
                }
            });
    }, []);

    const handleChange = async (file: any) => {
        await uploadFile(file);
    };
    const handleClick = (file: any, type: string) => {
        const documentbody = file.documentbody;
        const contentType = file.mimetype;
        const blob = b64toBlob(documentbody, contentType, 512);

        setBase64(documentbody);

        if (type === "preview") {
            const objectURL = URL.createObjectURL(blob);
            setFile(objectURL);
        } else {
            saveAs(blob, file.filename);
        }
    };
    const uploadFile = async (file: any) => {
        const base64 = await toBase64(file);
        createNote(base64, file.name, file.type, file.size);
    }
    const toBase64 = (file: Blob) => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
    const createNote = (docBody: any, fName: string, mType: string, fSize: number) => {
        var entity: any = {
            "documentbody": docBody.split(",")[1],
            "filename": fName,
            "mimetype": mType,
            "subject": "",
            "notetext": "",
            "isdocument": true,
        };
        props.componentContext.utils.getEntityMetadata(logicalName, "")
            .then(function (result: any) {
                entity[`objectid_${logicalName}@odata.bind`] = `/${result.EntitySetName}(${entityId})`;
            }, function (error: any) {
                console.log(error);
            });
        props.componentContext.webAPI.createRecord("annotation", entity)
            .then((result: any) => {
                entity["filesize"] = fSize;
                setAttachments([...attachments, entity]);
            })
            .catch((error: any) => {
                console.log("This is error: ", error)
            });
    };
    const b64toBlob = (b64Data: any, contentType: any, sliceSize: any) => {
        contentType = contentType || '';
        sliceSize = sliceSize || 512;

        var byteCharacters = atob(b64Data);
        var byteArrays = [];

        for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
            var slice = byteCharacters.slice(offset, offset + sliceSize);

            var byteNumbers = new Array(slice.length);
            for (var i = 0; i < slice.length; i++) {
                byteNumbers[i] = slice.charCodeAt(i);
            }

            var byteArray = new Uint8Array(byteNumbers);
            byteArrays.push(byteArray);
        }

        var blob = new Blob(byteArrays, { type: contentType });
        return blob;
    };
    return (
        <>
            <Row className='file-viewer-container'>
                <Col md={5} className="col file-viewer">
                    <div className='file-upload'>
                        <FileUploader
                            handleChange={handleChange}
                            name="file"
                            label="Upload file"
                        />
                    </div>
                    <div className="table-responsive table-container">
                        <table className='table'>
                            <thead>
                                <tr>
                                    <th>Filename</th>
                                    <th>Created on</th>
                                    <th>Size</th>
                                    <th className='download-th'>Download</th>
                                </tr>
                            </thead>
                            <tbody>
                                {(attachments.length > 0) && attachments.map((curr: any, ind: React.Key | null | undefined) => {
                                    const ext = "." + curr.filename.split(".")[1];
                                    const contenttype = curr.mimetype.split(".")[0];
                                    return (
                                        <tr key={ind}>
                                            <td onClick={() => handleClick(curr, "preview")}><span id={ext}>{curr.filename}</span></td>
                                            <td><span className="">{curr.createdon}</span></td>
                                            <td><span className="text-nowrap">{Math.round(curr.filesize * 0.001)} KBs</span></td>
                                            <td className='text-center'>
                                                <button className="download-btn" onClick={() => handleClick(curr, "download")}>

                                                </button>
                                            </td>
                                        </tr>
                                    )
                                })}
                            </tbody>
                        </table>
                    </div>
                </Col>
                <Col md={7} className="col viewer">
                    <div className="img-viewer">
                        <iframe src={file}></iframe>
                    </div>
                </Col>
            </Row>
        </>
    );
};
export default DocumentAttachment;