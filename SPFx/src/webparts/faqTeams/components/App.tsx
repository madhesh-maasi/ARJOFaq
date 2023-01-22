/* eslint-disable react/jsx-no-target-blank */
/* eslint-disable react-hooks/exhaustive-deps */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-debugger */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable prefer-const */
/* eslint-disable react/jsx-key */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable react/jsx-no-bind */
import * as React from "react";
import { useState, useEffect, useRef, useCallback } from "react";
import { ThemeProvider, PartialTheme } from "@fluentui/react";
import { makeStyles } from "@material-ui/core/styles";
import Accordion from "@material-ui/core/Accordion";
import AccordionDetails from "@material-ui/core/AccordionDetails";
import AccordionSummary from "@material-ui/core/AccordionSummary";
import Typography from "@material-ui/core/Typography";
import AddIcon from "@material-ui/icons/Add";
import RemoveIcon from "@material-ui/icons/Remove";
import SearchIcon from "@material-ui/icons/Search";
import Checkbox from "@material-ui/core/Checkbox";
import TextField from "@material-ui/core/TextField";
import Autocomplete from "@material-ui/lab/Autocomplete";
import CheckBoxOutlineBlankIcon from "@material-ui/icons/CheckBoxOutlineBlank";
import CheckBoxIcon from "@material-ui/icons/CheckBox";
import FormControl from "@material-ui/core/FormControl";
import { InputLabel, OutlinedInput, Select } from "@material-ui/core";
import Button from "@material-ui/core/Button";
import styles from "./App.module.scss";
import Modal from "@material-ui/core/Modal";
import Backdrop from "@material-ui/core/Backdrop";
import Fade from "@material-ui/core/Fade";
import { Web } from "@pnp/sp/webs";
// import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import ReactQuill, { Quill } from "react-quill";
import "react-quill/dist/quill.snow.css";
import ImageResize from "quill-image-resize-module-react";
import { Attachment } from "@pnp/sp/attachments";
Quill.register("modules/imageResize", ImageResize);

export interface IApp {
  tenantURL: string;
  siteName: string;
  teamName: string;
  channelName: string;
  // teamsContext: any;
  // userContext: any;
}

let arrAllQA = [];
let arrCategory = [];
let arrSubCategory = [];
let arrSelectedCategory = [];
let arrAttachments = [];
let strSearch = "";
let objModalSelected = {
  Category: 0,
  SubCategory: 0,
  Tag: "",
  Question: "",
  Answer: "",
  Link: "",
  Files: [],
};
let objModalError = {
  isCatError: false,
  isQuestionError: false,
  isLinkError: false,
};
let mainList = "";
let department = "";
const appTheme: PartialTheme = {
  palette: {
    themePrimary: "#015174",
  },
};
const useStyles = makeStyles((theme) => ({
  customTextField: {
    "& input::placeholder": {
      color: "#777379",
    },
  },
  modal: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  paper: {
    backgroundColor: theme.palette.background.paper,
    boxShadow: theme.shadows[5],
    padding: theme.spacing(2, 4, 3),
  },
}));
const bannerImg = require("../../../ExternalRef/IMG/bannerImg.png");
const sampleImg = require("../../../ExternalRef/IMG/SampleImage.jpg");
const uploadIcon = require("../../../ExternalRef/Icon/uploadIcon.png");
const icon = <CheckBoxOutlineBlankIcon color="primary" fontSize="small" />;
const checkedIcon = <CheckBoxIcon color="primary" fontSize="small" />;
// const spWeb = props.siteUrl;
const modules = {
  imageResize: {
    parchment: Quill.import("parchment"),
    modules: ["Resize", "DisplaySize"],
  },
  toolbar: [
    [
      {
        header: [1, 2, 3, 4, 5, 6, false],
      },
    ],
    ["link", "image"],
    ["bold", "italic", "underline"],
    [
      {
        color: [],
      },
      {
        background: [],
      },
    ],
    [
      {
        list: "ordered",
      },
      {
        list: "bullet",
      },
      {
        indent: "-1",
      },
      {
        indent: "+1",
      },
    ],
    ["clean"],
  ],
};
const formats = [
  "header",
  "bold",
  "italic",
  "underline",
  "list",
  "bullet",
  "indent",
  "background",
  "color",
  "image",
];
let toStateFiles = [];
export const App: React.FunctionComponent<IApp> = (props: IApp) => {
  // development  URLs
  // const tenantURL = "https://chandrudemo.sharepoint.com";
  // const spWeb = `${tenantURL}/sites/ARJOFAQ`;
  // let currentSite = "ARJOFAQ";
  // production URLs
  const tenantURL = props.tenantURL;
  const spWeb = `${tenantURL}/sites/${props.siteName}`;
  let currentSite = props.siteName;
  const web = Web(spWeb);
  const classes = useStyles();
  const [expanded, setExpanded] = useState("");
  const [accordianItem, setAccordianItem] = useState(arrAllQA);
  const [choicesCategory, setChoicesCategory] = useState(arrCategory);
  const [choicesSubCategory, setChoicesSubCategory] = useState([]);
  const [selectedCategory, setSelectedCategory] = useState(arrSelectedCategory);
  const [selectedSubCategory, setSelectedSubCategory] = useState({});
  const [seacrhValue, setSearchValue] = useState(strSearch);
  const [open, setOpen] = useState(false);
  const [modalError, setModalError] = useState(objModalError);
  const [editorState, setEditorState] = useState("");
  const [selectedFiles, setSelectedFiles] = useState([]);
  //  Get fil from Modal Inputs
  const getFiles = (e) => {
    let tempArrAttachments = e.target.files;
    for (let i = 0; i < tempArrAttachments.length; i++) {
      arrAttachments.push(tempArrAttachments[i]);
    }

    console.log(arrAttachments);
    setSelectedFiles([...arrAttachments]);
  };
  // Delete File
  const fileDelete = (fileName) => {
    arrAttachments = arrAttachments.filter((file) => file.name !== fileName);
    console.log(arrAttachments);
    setSelectedFiles([...arrAttachments]);
  };
  // Modal Control
  const handleOpen = () => {
    setOpen(true);
  };

  const handleClose = () => {
    setOpen(false);
  };

  // Modal Contro;
  const handleChange = (panel) => {
    setExpanded(panel !== expanded ? panel : "");
  };

  const handleFilter = (category, search) => {
    let filteredArr = [];
    filteredArr =
      category.length > 0
        ? arrAllQA.filter(
            (li) => li.Category && category.indexOf(li.Category.Title) >= 0
          )
        : arrAllQA;
    search !== ""
      ? (filteredArr = filteredArr.filter(
          (li) =>
            li.Question.toLowerCase().includes(search.toLowerCase()) ||
            // li.Answer && li.Answer.replace(/<[^>]+>/g, "")
            (li.Answer &&
              li.Answer.replace(/<[^>]+>/g, "")
                .toLowerCase()
                .includes(search.toLowerCase()))
        ))
      : filteredArr;
    category.length > 0 || search !== ""
      ? setAccordianItem([...filteredArr])
      : setAccordianItem([...arrAllQA]);
  };

  // Get Faq all data
  const getFAQ = (web) => {
    // Production
    // let team = props.teamName;
    // let Channel = props.channelName;
    // Dev
    let team = "HR Department";
    let Channel = "General";
    console.log(team);
    console.log(Channel);
    let arrFiles = [];
    web.lists
      .getByTitle("Config")
      .items.select(
        "*,PermissionFor/EMail,PermissionFor/Id,PermissionFor/Title,PermissionFor/Name"
      )
      .expand("PermissionFor")
      .get()
      .then(async (res) => {
        console.log(res);
        mainList = await res.filter(
          (li) => li.Team === team && li.Channel === Channel
        )[0].ListName;
        // Changes need to be done Department to List name
        await web
          .getFolderByServerRelativeUrl(
            `FAQ_Assets/${
              res.filter((li) => li.Team === team && li.Channel === Channel)[0]
                .ListName
            }`
          )
          .folders.get()
          .then((folders) => {
            console.log(folders);
            folders.forEach(async (folder) => {
              web
                .getFolderByServerRelativeUrl(
                  `FAQ_Assets/${
                    res.filter(
                      (li) => li.Team === team && li.Channel === Channel
                    )[0].ListName
                  }/${folder.Name}`
                )
                .files.get()
                .then((files) => {
                  for (let i = 0; i < files.length; i++) {
                    let _ServerRelativeUrl = files[i].ServerRelativeUrl;
                    web
                      .getFileByServerRelativeUrl(_ServerRelativeUrl)
                      .getItem()
                      .then(async (item) => {
                        await arrFiles.push({
                          FileName: files[i].Name,
                          Url: `${tenantURL}${files[i].ServerRelativeUrl}`,
                          ListID: parseInt(folder.Name),
                        });
                        console.log(arrFiles);
                      })
                      .then(async () => {
                        // Getting Data from list
                        await web.lists
                          .getByTitle(
                            res.filter(
                              (li) => li.Team === team && li.Channel === Channel
                            )[0].ListName
                          )
                          .items.select(
                            "*,Category/Title,Category/ID,SubCategory/Title,SubCategory/ID"
                          )
                          .expand("Category,SubCategory")
                          .get()
                          .then((res) => {
                            res = res.filter((li) => li.Approved);
                            arrAllQA = res.map((li) => {
                              return {
                                Answer: li.Answer,
                                Approved: li.Approved,
                                Category: li.Category,
                                ID: li.ID,
                                Question: li.Question,
                                SubCategory: li.SubCategory,
                                Files: arrFiles.filter(
                                  (file) => file.ListID === li.ID
                                ),
                              };
                            });
                            console.log(arrAllQA);
                            setAccordianItem([...arrAllQA]);
                          })
                          .catch((error) => console.log(error));
                      });
                  }
                });
            });
          });
      });
  };
  const getCategory = (web) => {
    // Category
    web.lists
      .getByTitle("Category")
      .items.get()
      .then((res) => {
        arrCategory = res.map((li) => ({ title: li.Title, id: li.ID }));
        setChoicesCategory([...arrCategory]);
      })
      .catch((error) => console.log(error));
  };
  const getSubCategory = (web) => {
    web.lists
      .getByTitle("SubCategory")
      .items.select("*,Category/Title,Category/ID")
      .expand("Category")
      .get()
      .then((res) => {
        arrSubCategory = res.map((li) => ({
          title: li.Title,
          category: li.Category.Title,
          id: li.ID,
        }));
        console.log(arrSubCategory);

        // setChoicesSubCategory([...arrSubCategory]);
      })
      .catch((error) => console.log(error));
  };
  // Bug in Resetting Sub Category
  const resetSubCategory = () => {
    objModalSelected.SubCategory = 0;
    let objSelectedSub =
      objModalSelected.Category !== 0
        ? arrSubCategory.filter(
            (li) =>
              li.category ===
              arrCategory.filter(
                (item) => item.id === objModalSelected.Category
              )[0].title
          )
        : [];
    setSelectedSubCategory({ ...{} });
    setChoicesSubCategory([...objSelectedSub]);
  };
  const resetError = () => {
    objModalError = {
      isCatError: false,
      isQuestionError: false,
      isLinkError: false,
    };
    setModalError({ ...objModalError });
  };
  const handleSubmitModal = (obj) => {
    console.log(obj);
    let regex =
      /(?:https?):\/\/(\w+:?\w*)?(\S+)(:\d+)?(\/|\/([\w#!:.?+=&%!\-\/]))?/;
    if (obj.Category === 0) {
      resetError();
      objModalError.isCatError = true;
      setModalError({ ...objModalError });

      return false;
    } else if (obj.Question === "") {
      resetError();
      objModalError.isQuestionError = true;
      setModalError({ ...objModalError });
      return false;
    } else if (!regex.test(obj.Link) && obj.Link !== "") {
      resetError();
      objModalError.isLinkError = true;
      setModalError({ ...objModalError });

      return false;
    } else {
      resetError();
      let addItem = {
        Question: obj.Question,
        Answer: obj.Answer,
        CategoryId: obj.Category,
        SubCategoryId: obj.SubCategory !== 0 ? obj.SubCategory : null,
        Tag: obj.Tag,
        Link: {
          Description: obj.Link !== "" ? "Click here" : "",
          Url: obj.Link !== "" ? obj.Link : "",
        },
      };
      console.log(addItem);
      web.lists
        .getByTitle(mainList)
        .items.add(addItem)
        .then((res) => {
          console.log(res);
          let uploadID = res.data.ID;
          let updateItem = {
            LinkToFolder: {
              Description: obj.Link !== "" ? "Click here" : "",
              Url: `${spWeb}/FAQ_Assets/Forms/AllItems.aspx?id=/sites/${currentSite}/FAQ_Assets/${mainList}/${uploadID}`,
            },
          };
          web.folders.add(`FAQ_Assets/${mainList}/${uploadID}`).then((res) => {
            console.log(res);
            web.lists
              .getByTitle(mainList)
              .items.getById(uploadID)
              .update(updateItem);
            if (selectedFiles.length > 0) {
              selectedFiles.forEach((file) => {
                if (file !== undefined || file !== null) {
                  //assuming that the name of document library is Documents, change as per your requirement,
                  //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
                  web
                    .getFolderByServerRelativeUrl(
                      `FAQ_Assets/${mainList}/${uploadID}`
                    )
                    .files.add(file.name, file, true)
                    .then((data) =>
                      data.file.getItem().then((fileItem: any) => {
                        fileItem
                          .update({
                            ListID: uploadID,
                          })
                          .then((updatedItem: any) => {
                            console.log(updatedItem);
                          })
                          .catch((error: any) => {
                            console.log(error);
                          });
                      })
                    );
                }
              });
            }
            setEditorState("");
            arrAttachments = [];
            setSelectedFiles([]);
            setSelectedSubCategory({});
            handleClose();
          });
        });
    }
  };

  useEffect(() => {
    getFAQ(web);
    getCategory(web);
    getSubCategory(web);
  }, []);

  return (
    <ThemeProvider theme={appTheme}>
      <div className={styles.App}>
        <div
          className={styles.bannerSection}
          style={{ background: `url(${bannerImg})` }}
        >
          <div className={styles.bannerTitleSection}>
            <div className={styles.bannerTitle}>Frequently Asked Question</div>
          </div>
        </div>
        {/* Filter section */}
        <div className={styles.FilterSection}>
          <Autocomplete
            className="filterSelect"
            size="small"
            multiple
            id="checkboxes-tags-demo"
            options={choicesCategory}
            disableCloseOnSelect
            onChange={(e, val) => {
              arrSelectedCategory = val.map((row) => row.title);
              setSelectedCategory([...arrSelectedCategory]);
              handleFilter(arrSelectedCategory, seacrhValue);
            }}
            // eslint-disable-next-line react/jsx-no-bind
            getOptionLabel={(option) => option.title}
            renderOption={(option, { selected }) => (
              <React.Fragment>
                <Checkbox
                  icon={icon}
                  checkedIcon={checkedIcon}
                  style={{ marginRight: 8 }}
                  checked={selected}
                />
                {option.title}
              </React.Fragment>
            )}
            style={{ width: 200, marginRight: 24 }}
            renderInput={(params) => (
              <TextField
                {...params}
                variant="outlined"
                // label="Checkboxes"
                placeholder="Select a category"
              />
            )}
          />
          <FormControl
            style={{ width: 200 }}
            variant="outlined"
            size="small"
            className="filterSearch"
          >
            <OutlinedInput
              onChange={(event) => {
                strSearch = event.target.value;
                setSearchValue("");
                setSearchValue(strSearch);
                handleFilter(selectedCategory, strSearch);
              }}
              inputProps={{ shrink: false }}
              placeholder="Search"
              id="outlined-adornment-amount"
              // value={values.amount}
              // onChange={handleChange('amount')}
              startAdornment={<SearchIcon />}
              labelWidth={60}
            />
          </FormControl>
          <Button
            onClick={handleOpen}
            variant="contained"
            style={{
              background: appTheme.palette.themePrimary,
              color: "#fff",
              marginTop: -5,
              marginLeft: 24,
            }}
          >
            New Entry
          </Button>
        </div>
        {/* Filter section */}
        {/* Accordian Section */}
        <div className={styles.AccordionSection}>
          {accordianItem.length > 0 ? (
            accordianItem.map((item, i) => {
              return (
                <Accordion
                  className={styles.AccordionItem}
                  expanded={expanded === `panel${i}`}
                  onChange={() => handleChange(`panel${i}`)}
                >
                  <AccordionSummary
                    expandIcon={
                      expanded === `panel${i}` ? <RemoveIcon /> : <AddIcon />
                    }
                    aria-label="Expand"
                    aria-controls="additional-actions1-content"
                    id="additional-actions1-header"
                  >
                    <Typography
                      variant="h6"
                      component="h2"
                      className={styles.AccordionTitle}
                    >
                      {item.Question}
                    </Typography>
                  </AccordionSummary>
                  <AccordionDetails className={styles.AnswerSection}>
                    <div className={styles.AccDetails}>
                      <div className={styles.Divider}></div>
                      <div className={styles.AnswerWithImage}>
                        <div
                          dangerouslySetInnerHTML={{
                            __html: item.Answer,
                          }}
                        />
                      </div>

                      <div className={styles.fileAttached}>
                        {item.Files.length > 0 && (
                          <div className={styles.attachmentTitle}>
                            Attachments:
                          </div>
                        )}
                        {item.Files.length > 0 &&
                          item.Files.map((file) => {
                            return (
                              <a
                                href={`${file.Url}`}
                                target="_blank"
                                data-interception="off"
                                style={{ margin: 6 }}
                              >
                                {file.FileName}
                              </a>
                            );
                          })}
                      </div>
                    </div>
                  </AccordionDetails>
                </Accordion>
              );
            })
          ) : (
            <div className={styles.DataNotFound}>No Input Found</div>
          )}
        </div>
        {/* Accordian Section */}
      </div>
      {/*  Modal Section */}
      <Modal
        aria-labelledby="transition-modal-title"
        aria-describedby="transition-modal-description"
        className={classes.modal}
        open={open}
        onClose={handleClose}
        closeAfterTransition
        BackdropComponent={Backdrop}
        BackdropProps={{
          timeout: 500,
        }}
      >
        <Fade
          in={open}
          style={{ maxHeight: "100vh", maxWidth: "720px", overflowY: "auto" }}
        >
          <div className={classes.paper}>
            <h4
              style={{
                // margin: 16,
                color: appTheme.palette.themePrimary,
                textAlign: "center",
                fontWeight: 600,
                fontSize: 18,
              }}
            >
              New Question and Answer
            </h4>

            <div className={styles.modalInputSection}>
              <div className={styles.modalInput} style={{ marginRight: 16 }}>
                <Autocomplete
                  size="small"
                  id="checkboxes-tags-demo"
                  options={choicesCategory}
                  onChange={(e, val) => {
                    val
                      ? (objModalSelected.Category = val.id)
                      : (objModalSelected.Category = 0);
                    resetSubCategory();
                  }}
                  // eslint-disable-next-line react/jsx-no-bind
                  getOptionLabel={(option) => option.title}
                  renderOption={(option, { selected }) => (
                    <React.Fragment>{option.title}</React.Fragment>
                  )}
                  style={{ width: 320 }}
                  renderInput={(params) => (
                    <TextField
                      error={modalError.isCatError}
                      helperText={
                        modalError.isCatError ? "Please Select Category" : ""
                      }
                      {...params}
                      variant="outlined"
                      label="Category"
                      placeholder="Select a category"
                    />
                  )}
                />
              </div>
              <div className={styles.modalInput}>
                <Autocomplete
                  size="small"
                  id="checkboxes-tags-demo"
                  options={choicesSubCategory}
                  value={selectedSubCategory}
                  onChange={(e, val) => {
                    val && val.id !== 0
                      ? (objModalSelected.SubCategory = val.id)
                      : (objModalSelected.SubCategory = 0);
                    let objSub = choicesSubCategory.filter(
                      (li) => li.id === objModalSelected.SubCategory
                    );
                    setSelectedSubCategory({
                      ...(objSub.length > 0 ? objSub[0] : {}),
                    });
                    console.log(objModalSelected);
                  }}
                  // eslint-disable-next-line react/jsx-no-bind
                  getOptionLabel={(option) => option.title}
                  renderOption={(option, { selected }) => (
                    <React.Fragment>{option.title}</React.Fragment>
                  )}
                  style={{ width: 320 }}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      variant="outlined"
                      label="Sub Category"
                      placeholder="Select a Sub Category"
                    />
                  )}
                />
              </div>
            </div>
            <div className={styles.modalInputSection}>
              <div className={styles.modalInput} style={{ marginRight: 16 }}>
                <TextField
                  size="small"
                  style={{ width: 320 }}
                  id="outlined-basic"
                  label="Tag"
                  variant="outlined"
                  onChange={(e) => {
                    objModalSelected.Tag = e.target.value;
                  }}
                />
              </div>
              <div className={styles.modalInput}>
                <TextField
                  size="small"
                  error={modalError.isLinkError}
                  helperText={
                    modalError.isLinkError ? "Please enter valid URL" : ""
                  }
                  style={{ width: 320 }}
                  id="outlined-basic"
                  label="Link"
                  variant="outlined"
                  onChange={(e) => {
                    objModalSelected.Link = e.target.value;
                  }}
                />
              </div>
            </div>

            <div className={styles.modalInput}>
              <TextField
                error={modalError.isQuestionError}
                helperText={
                  modalError.isQuestionError ? "Please enter question" : ""
                }
                style={{ width: "100%" }}
                size="small"
                id="outlined-basic"
                label="Question"
                variant="outlined"
                onChange={(e) => {
                  objModalSelected.Question = e.target.value;
                }}
              />
            </div>
            <div className={styles.modalInput}>
              {/* <Button variant="contained" component="label">
                Upload File
                <input
                  type="file"
                  multiple
                  onChange={(e) => {
                    console.log(e.target.files);
                  }}
                /> 
              </Button> */}
              <div className={styles.customUpload}>
                <TextField
                  onChange={(e) => getFiles(e)}
                  type="file"
                  style={{ marginRight: "20px", display: "none" }}
                  // className={styles.inputsm}
                  inputProps={{
                    multiple: true,
                  }}
                  id="customUpload"
                />
                <label
                  htmlFor="customUpload"
                  style={{
                    color: "#000",
                    display: "block",
                    cursor: "pointer",
                    width: "100%",
                  }}
                >
                  Select file to upload
                </label>
                <img
                  className={styles.customUploadIcon}
                  src={`${uploadIcon}`}
                />
              </div>

              <div className={styles.FileNames}>
                {selectedFiles.length > 0 &&
                  selectedFiles.map((file) => {
                    return (
                      <div className={styles.fileSection}>
                        <span>{file.name}</span>
                        <span
                          onClick={() => fileDelete(file.name)}
                          style={{
                            marginLeft: 6,
                            color: "red",
                            cursor: "pointer",
                          }}
                        >
                          X
                        </span>
                      </div>
                    );
                  })}
              </div>
            </div>
            <div className={styles.modalInput}>
              <ReactQuill
                placeholder="Please enter answer . . . "
                theme="snow"
                modules={modules}
                formats={formats}
                value={editorState}
                onChange={(e) => {
                  console.log(e);
                  objModalSelected.Answer = e;
                  console.log(objModalSelected);
                  setEditorState(objModalSelected.Answer);
                }}
                style={{
                  height: "auto",
                }}
              />
            </div>
            <div className={styles.modalBtnSection}>
              <Button
                onClick={() => {
                  handleClose();
                  resetError();
                }}
                variant="contained"
              >
                Cancel
              </Button>
              <Button
                onClick={() => {
                  handleSubmitModal(objModalSelected);
                }}
                variant="contained"
                color="primary"
                style={{
                  marginLeft: 8,
                  background: appTheme.palette.themePrimary,
                }}
              >
                Submit
              </Button>
            </div>
          </div>
        </Fade>
      </Modal>
      {/* Modal Section */}
    </ThemeProvider>
  );
};
