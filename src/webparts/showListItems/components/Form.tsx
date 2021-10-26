import * as React from "react"

function Form() {
    return (
        <div>
            <form>
                <label htmlFor="name"></label>
                <input type="text" id="name" />
                <label htmlFor="lastName"></label>
                <input type="text" id="lastName" />
                <input type="submit" value="Salvar" />
            </form>
        </div>
    )
}

export default Form