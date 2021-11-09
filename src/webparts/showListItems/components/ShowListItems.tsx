import * as React from "react";
import { IShowListItemsProps } from "./IShowListItemsProps";
import { Route, Switch, HashRouter } from "react-router-dom";

import Home from "../pages/Home";
import Cadastro from "../pages/Cadastro";
import Navbar from "./layout/Navbar";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { FaEdit, FaTrashAlt } from "react-icons/fa";
import "bootstrap/dist/css/bootstrap.min.css";
import "./ShowListItems.module.css";

export interface getListItems {
  Items: any;
}

export default class ShowListItems extends React.Component<IShowListItemsProps, getListItems, {} > {
  constructor(props) {
    super(props)
    this.state = {
      Items: [],
    }
  }

  async componentDidMount() {
    // const items: any[] = await sp.web.lists.getByTitle(this.props.listName).items.get()
    const items: any[] = await sp.web.lists.getByTitle("Pessoas").items.get();

    this.setState({
      Items: items,
    });
  }

  async itemDelete(id) {
    console.log(id)
    console.log(typeof id)
    //await sp.web.lists.getByTitle("Pessoas").items.getById(id).delete();
  }

  tHeadData = ["#", "Nome", "Sobrenome", "E-mail", "Gênero", "Ações"];

  renderTbody(value, index) {
    return (
      <tr>
        <td>{index + 1}</td>
        <td>{value.Title}</td>
        <td>{value.Sobrenome}</td>
        <td>{value.Email}</td>
        <td>{value.Genero}</td>
        <td className="btnsAction">
          <a onClick={ this.itemDelete(value.Id)} title={value.Id}>
            <FaEdit className="btnEdit" />
          </a>
          &nbsp;
          <a href="">
            <FaTrashAlt className="btnDelete" />
          </a>
        </td>
      </tr>
    );
  }

  public render(): React.ReactElement<IShowListItemsProps> {
    return (
      <div className="container">
        <HashRouter>

          <Navbar />

          <Switch>

            <Route exact path="/">
              <Home
                thead={this.tHeadData}
                tbody={this.state.Items.map(this.renderTbody)}
              />
            </Route>

            <Route path="/cadastro">
              <Cadastro />
            </Route>

          </Switch>

        </HashRouter>
      </div>
    );
  }
}
