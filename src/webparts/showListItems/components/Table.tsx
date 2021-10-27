import * as React from "react"

import "bootstrap/dist/css/bootstrap.min.css"

function Table({data}) {
    return (
        <div>
            <table className="table table-striped">
            <thead>
              <th>Nome</th>
              <th>Sobrenome</th>
              <th>Idade</th>
              <th>Gênero</th>
            </thead>
            <tbody>
              {data}
            </tbody>
          </table>
        </div>
    )
}

export default Table