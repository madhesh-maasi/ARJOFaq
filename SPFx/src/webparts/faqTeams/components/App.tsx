/* eslint-disable prefer-const */
/* eslint-disable react/jsx-key */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable react/jsx-no-bind */
import * as React from "react";
import { useState, useEffect, useRef } from "react";
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
import { OutlinedInput } from "@material-ui/core";
import Button from "@material-ui/core/Button";
import styles from "./App.module.scss";
import Modal from "@material-ui/core/Modal";
import Backdrop from "@material-ui/core/Backdrop";
import Fade from "@material-ui/core/Fade";
import { Web } from "@pnp/sp/webs";
export interface IApp {
  siteUrl: string;
}

let arrAllQA = [];
let arrCategory = [];
let arrSubCategory = [];
let arrSelectedCategory = [];
let strSearch = "";
const appTheme: PartialTheme = {
  palette: {
    themePrimary: "#015174",
  },
};
const useStyles = makeStyles((theme) => ({
  modal: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  paper: {
    backgroundColor: theme.palette.background.paper,
    border: "2px solid #000",
    boxShadow: theme.shadows[5],
    padding: theme.spacing(2, 4, 3),
  },
}));
const bannerImg = require("../../../ExternalRef/IMG/bannerImg.png");
const sampleImg = require("../../../ExternalRef/IMG/SampleImage.jpg");
const icon = <CheckBoxOutlineBlankIcon color="primary" fontSize="small" />;
const checkedIcon = <CheckBoxIcon color="primary" fontSize="small" />;

export const App: React.FunctionComponent<IApp> = (props: IApp) => {
  const classes = useStyles();
  const [expanded, setExpanded] = useState("");
  const [accordianItem, setAccordianItem] = useState(arrAllQA);
  const [choicesCategory, setChoicesCategory] = useState(arrCategory);
  const [choicesSubCategory, setChoicesSubCategory] = useState(arrSubCategory);
  const [selectedCategory, setSelectedCategory] = useState(arrSelectedCategory);
  const [seacrhValue, setSearchValue] = useState(strSearch);
  const [open, setOpen] = React.useState(false);
  // Modal Contro;
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
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    web.lists
      .getByTitle("FAQ")
      .items.select(
        "*,Category/Title,Category/ID,SubCategory/Title,SubCategory/ID"
      )
      .expand("Category,SubCategory")
      .get()
      .then((res) => {
        arrAllQA = res;
        console.log(arrAllQA);
        setAccordianItem([...arrAllQA]);
      })
      .catch((error) => console.log(error));
  };
  const getCategory = (web) => {
    // Category
    web.lists
      .getByTitle("Category")
      .items.get()
      .then((res) => {
        arrCategory = res.map((li) => ({ title: li.Title }));
        setChoicesCategory([...arrCategory]);
      })
      .catch((error) => console.log(error));
  };
  const getSubCategory = (web) => {};
  useEffect(() => {
    const spWeb = "https://chandrudemo.sharepoint.com/sites/ARJOFAQ";
    // const spWeb = props.siteUrl;
    const web = Web(spWeb);
    getFAQ(web);
    getCategory(web);
  }, [props.siteUrl]);
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
              console.log(arrSelectedCategory);
              setSelectedCategory([...arrSelectedCategory]);
              console.log(selectedCategory);
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
            style={{ width: 260, marginRight: 24 }}
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
            style={{ width: 260 }}
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
      {/* Modal Section */}
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
        <Fade in={open}>
          <div className={classes.paper}>
            <h4>New Question and Answer</h4>
            <Autocomplete
              className="filterSelect"
              size="small"
              id="checkboxes-tags-demo"
              options={choicesCategory}
              disableCloseOnSelect
              onChange={(e, val) => {
                console.log(val);
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
              style={{ width: 260, marginRight: 24 }}
              renderInput={(params) => (
                <TextField
                  {...params}
                  variant="outlined"
                  // label="Checkboxes"
                  placeholder="Select a category"
                />
              )}
            />
          </div>
        </Fade>
      </Modal>
      {/* Modal Section */}
    </ThemeProvider>
  );
};
