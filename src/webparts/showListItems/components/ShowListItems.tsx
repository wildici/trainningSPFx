import * as React from 'react'
import { IShowListItemsProps } from './IShowListItemsProps'
import { escape } from '@microsoft/sp-lodash-subset'
import Form from './Form'
import Table from './Table'

import { sp } from "@pnp/sp"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"

export interface getListItems{
  Items: any
}

export default class ShowListItems extends React.Component<IShowListItemsProps, getListItems, {}> {
  
  constructor(props) {
    super(props)
    this.state = {
      Items: []
    }
  }

  async componentDidMount() {
    // const items: any[] = await sp.web.lists.getByTitle(this.props.listName).items.get()
    const items: any[] = await sp.web.lists.getByTitle("Pessoas").items.get()
    this.setState({
      Items: items
    })
  }

  renderHtml(value) {
    return (
      <tr>
        <td>{value.Title}</td>
        <td>{value.Sobrenome}</td>
        <td>{value.Email}</td>
        <td>{value.Genero}</td>
      </tr>
    )
  }

  public render(): React.ReactElement<IShowListItemsProps> {

    return (
      <div>
          <Table data={this.state.Items.map(this.renderHtml)} />
      </div>
    )
  }

}