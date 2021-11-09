import * as React from "react"
import { Link } from 'react-router-dom'

import './Navbar.module.css'

function Navbar() {
    return (
        <>
            <ul className="list">
                <li className="item">
                    <Link to="/">Home</Link>
                </li>
                <li className="item">
                    <Link to="/cadastro">Cadastro</Link>
                </li>
            </ul>
        </>
    )
}

export default Navbar