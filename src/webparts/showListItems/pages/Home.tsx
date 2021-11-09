import * as React from "react"

import './Home.module.css'

function Home( {thead, tbody} ) {

  return (
    <div>
      <h2>Lista de Contatos</h2>
      <table className="table table-striped">
            <thead>
              <tr>
              { thead.map( e => ( <th>{e}</th> ) ) }
              </tr>
            </thead>
            <tbody>
              {tbody}
            </tbody>
          </table>
    </div>
  )
}

export default Home
