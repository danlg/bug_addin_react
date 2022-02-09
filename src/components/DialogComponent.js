import * as React from "react";
import PropTypes from "prop-types";
import {DefaultButton} from "@fluentui/react";
import Progress from "./Progress";

import {writeToWordImpl} from '../commands/commands.js'

export default class DialogComponent extends React.Component {
    constructor(props, context) {
        super(props, context)
        this.state = {
            listItems: [],
            displayText:""
        };
        this.count = 0
    }

    async componentDidMount() {
        console.log('componentDidMount');
    }

    writeToWord = async (text, evt) => {
        writeToWordImpl(text)
    }
    unit_test = async (text, evt) => {
        console.log("Hello UNIT TEST (11)")
        // see https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/dialog-api-in-office-add-ins.md
        writeToWordImpl(text, evt)
    }

    render() {
        const { title, isOfficeInitialized } = this.props;
        if (!isOfficeInitialized) {
          return (
            <Progress
              title={title}
              logo={require("../../assets/icon-128.png")}
              message="Please sideload your addin to see app body."
            />
          );
        }
        return ( 
        // <React.StrictMode >
                // <body class="ms-font-l">
                    <main className="ms-firstrun-instructionstep">
                        <div className="ms-firstrun-instructionstep__welcome-body">
                            <p align="center">
                                <DefaultButton 
                                    onClick={(e) => this.unit_test("UNIT TEST", e)}
                                    >UNIT TEST ***</DefaultButton>
                                </p>
                        </div>
                    </main>
                // </body>
        // </React.StrictMode>
        ) //end of return
    } //enf of render()
    
}

DialogComponent.propTypes = {
    title: PropTypes.string,
    isOfficeInitialized: PropTypes.bool,
    displayText: PropTypes.string,
};
