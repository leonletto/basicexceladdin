import * as React from 'react';
import {Button, ButtonType, TextField} from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, {HeroListItem} from './HeroList';
import Progress from './Progress';
import axios, {AxiosResponse} from 'axios';
import * as https from 'https';

const httpsAgent = new https.Agent({keepAlive: true, maxSockets: 40, maxFreeSockets: 10, timeout: 900000});


/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  currentPage?: string;
  apiKey: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      currentPage: 'home',
      apiKey: ''
    };
    this.handleChangeApiKey = this.handleChangeApiKey.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  handleChangeApiKey(event) {
    this.setState({apiKey: event.target.value});
  }

  handleSubmit(event) {
    event.preventDefault();
  }

  componentDidMount() {
    this.setState({
      listItems: [
        // {
        //   icon: 'Ribbon',
        //   primaryText: 'Achieve more with Office integration'
        // },
        // {
        //   icon: 'Unlock',
        //   primaryText: 'Unlock features and functionality'
        // },
        // {
        //   icon: 'Design',
        //   primaryText: 'Create and visualize like a pro'
        // }
        {
          icon: 'Ribbon',
          primaryText: 'Try out this Demo App with an API call'
        }
      ]
    });
  }


  apiKey = '';
  getNews = async (apikey: string) => {
    const columnLetters: string[] = [];
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    for (let x = 0; x < alphabet.length; x++) {
      for (let y = 0; y < alphabet.length; y++) {
        if (x === 0) {
          columnLetters.push(alphabet[y]);
        } else {
          columnLetters.push(alphabet[x - 1] + alphabet[y]);
        }
      }
    }

    await axios({
      url: 'https://' + window.location.host + '/newsapitopheadlines' + '?country=us&category=business&apikey=' + apikey,
      method: 'get',
      timeout: 90000,
      httpsAgent: httpsAgent,
      headers: {
        'Content-Type': `application/x-www-form-urlencoded`,
        Accept: 'application/json',
      },
    }).then(async (response: AxiosResponse) => {
      let models = [];

      models.push(['Source', 'Title', 'Description', 'PublishedAt', 'Content', 'Url', 'UrlToImage']);
      await response.data.articles.map(async (data) => {
        models.push([data.source.name, data.title, data.description, data.publishedAt, data.content, data.url, data.urlToImage]);
      });
      await Excel.run(async (context: Excel.RequestContext) => {
        let sheetTitle = 'TopNews';
        context.workbook.worksheets.getItemOrNullObject(sheetTitle).delete();
        const sheet = context.workbook.worksheets.add(sheetTitle);
        let dataWidth = models[0].length;
        let dataLength = models.length;
        for (let i = 0; i < dataLength; i + 1000 > dataLength ? i = dataLength : i = i + 1000) {
          var range = sheet.getRange('A' + (i + 1) + ':' + columnLetters[dataWidth - 1] + (i + 1000 > dataLength ? dataLength : i + 1000));
          range.values = models.slice(i, (i + 1000 > dataLength ? dataLength : i + 1000));
          sheet.activate();
          await context.sync();
        }

        sheet.getUsedRange().getEntireColumn().format.autofitColumns();
        sheet.getUsedRange().getEntireRow().format.autofitRows();
        sheet.activate();
      });
    });

  };

  click = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load('address');

        // Update the fill color
        range.format.fill.color = 'yellow';

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const {title, isOfficeInitialized} = this.props;

    // Set the body of the page based on state.
    let body;

    if (this.state.currentPage === 'home') {
      body = <HeroList message="Get some News!" items={this.state.listItems}>
        {/*<p className="ms-font-l">*/}
        {/*  Modify the source files, then click <b>Run</b>.*/}
        {/*</p>*/}
        {/*<Button*/}
        {/*  className="ms-welcome__action"*/}
        {/*  buttonType={ButtonType.hero}*/}
        {/*  iconProps={{iconName: 'ChevronRight'}}*/}
        {/*  onClick={this.click}*/}
        {/*>*/}
        {/*  Run*/}
        {/*</Button>*/}
        <form onSubmit={this.handleSubmit}>
          <TextField
            label="ApiKey"
            className='ms-welcome__formitems'
            autoCapitalize='none'
            type=''
            value={this.state.apiKey}
            onChange={this.handleChangeApiKey}
            placeholder='Enter NewsApi ApiKey'
          /> <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{iconName: 'ChevronRight'}}
          onClick={() => {
            this.getNews(this.state.apiKey);
          }}
        >
          Get News
        </Button>
        </form>

      </HeroList>;
    }

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body."/>
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome"/>
        {body}
        {/*<HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>*/}
        {/*  <p className="ms-font-l">*/}
        {/*    Modify the source files, then click <b>Run</b>.*/}
        {/*  </p>*/}
        {/*  <Button*/}
        {/*    className="ms-welcome__action"*/}
        {/*    buttonType={ButtonType.hero}*/}
        {/*    iconProps={{iconName: 'ChevronRight'}}*/}
        {/*    onClick={this.click}*/}
        {/*  >*/}
        {/*    Run*/}
        {/*  </Button>*/}
        {/*</HeroList>*/}
      </div>
    );
  }
}
