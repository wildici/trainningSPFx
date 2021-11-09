import * as React from "react"
import { useForm } from "react-hook-form"

import { sp } from "@pnp/sp"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"

import './Cadastro.module.css'

type Inputs = {
  Title: string
  Sobrenome: string
  Email: string
  Genero: string
}

function Cadastro() {
  const { register, watch, handleSubmit } = useForm<Inputs>()

  const saveData = (data) => {
    sp.web.lists
      .getByTitle("Pessoas")
      .items.add(data)
      .then((e) => {
        console.log("Sucesso " + e)
      })
      .catch((e) => {
        console.log("Erro " + e)
      })
  }

  return (
    <div>
      <h2>Painel de Cadastro</h2>
      <form onSubmit={handleSubmit(saveData)}>
        <div className="mb-3">
          <label htmlFor="name" className="form-label">
            Nome:
          </label>
          <input
            type="text"
            className="form-control"
            id="name"
            placeholder="Digite seu nome"
            ref={register}
            name="Title"
          />
        </div>
        <div className="mb-3">
          <label htmlFor="lastName" className="form-label">
            Sobrenome:
          </label>
          <input
            type="text"
            className="form-control"
            id="lastName"
            placeholder="Digite seu sobrenome"
            ref={register}
            name="Sobrenome"
          />
        </div>
        <div className="mb-3">
          <label htmlFor="email" className="form-label">
            E-mail:
          </label>
          <input
            type="email"
            className="form-control"
            id="email"
            placeholder="Digite seu e-mail"
            ref={register}
            name="Email"
          />
        </div>

        <button type="submit" className="btn btn-outline-primary">
          Salvar
        </button>
      </form>
    </div>
  );
}

export default Cadastro
