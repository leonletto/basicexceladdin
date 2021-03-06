import * as React from 'react';

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}


export default class Header extends React.Component<HeaderProps> {

  openDocs = async () => {
    window.open('https://' + window.location.host);
  };

  render() {
    const {title, logo, message} = this.props;


    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500"
               style={{height: 100, paddingBottom: 50}}>
        <img onClick={this.openDocs} width="90" height="90" src={logo} alt={title} title={title}/>
        <h1 className="ms-fontSize-xlPlus ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
      </section>
    );
  }
}
