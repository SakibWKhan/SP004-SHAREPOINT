
import * as React from 'react';

 

import styles from './Sales.module.scss';

 

import { ISalesProps } from './ISalesProps';

 

import { escape } from '@microsoft/sp-lodash-subset';

 

import { SPHttpClient, SPHttpClientResponse ,SPHttpClientConfiguration, ISPHttpClientOptions} from '@microsoft/sp-http';

 

import { Dropdown } from 'office-ui-fabric-react';

 

 

 

export interface ICustomer {

 

 

 

 

 

 

 

  CustomerID: number;

 

 

 

 

  CustomerName: string;

 

 

 

 

 

 

 

}

 

 

 

 

 

 

 

export interface IProduct {

 

 

 

 

 

 

 

  ProductID: number;



 

 

 

  ProductName: string;

 

 

 

 

ProductType: string;

 

 

 

 

  ProductExpireDate: string;

 

 

 

 

  ProductUnitPrice: number;

 

 

 

 

 

 

 

}

 

 

 

 

 

 

 

export interface IOrder {

 

 

 

 

 

 

 

  ID: number,

 

 

 

 

  OrderID: string;

 

 

 

 

  CustomerID: string;

 

 

 

 

  ProductID: string;

 

 

 

 

  UnitsSold: number;

 

 

 

 

  UnitPrice: number;

 

 

 

 

  SaleValue: number;

 

 

 

 

  OrderStatus: string;

 

 

 

 

 

 

 

}

 

 

 

 

 

 

 

export interface IFormState {

 

 

 

 

  customers: ICustomer[];

 

 

 

 

  products: IProduct[];

 

 

 

 

  orders: IOrder[];

 

 

 

 

  selectedCustomer: string;

 

 

 

 

  selectedProduct: string;

 

 

 

 

  selectedProductType: string;

 

 

 

 

  selectedExpiryDate: string;

 

 

 

 

  selectedUnitPrice: number;

 

 

 

 

  unitsSold: number;

 

 

 

 

  selectedOrderID: string;

 

 

 

 

 

 

 

}

 

 

 

 

 

 

 

export default class Form extends React.Component<ISalesProps, IFormState> {

 

 

 

 

 

 

 

  constructor(props: ISalesProps) {

 

 

 

 

    super(props);

 

 

 

 

    this.state = {

 

 

 

 

      customers: [],

 

 

 

 

      products: [],

 

 

 

 

      orders: [],

 

 

 

 

      selectedCustomer: '',

 

 

 

 

      selectedProduct: '',

 

 

 

 

      selectedProductType: '',

 

 

 

 

      selectedExpiryDate: '',

 

 

 

 

      selectedUnitPrice: 0,

 

 

 

 

      unitsSold: 0,

 

 

 

 

      selectedOrderID: ''

 

 

 

 

    };

 

 

 

 

  }

 

 

 

 

  public componentDidMount(): void {

 

 

 

 

    this.loadCustomers();

 

 

 

 

    this.loadProducts();

 

 

 

 

    this.loadOrders();

 

 

 

 

  }

 

 

 

 

  // Inside the loadCustomers() method:

 

 

 

 

 

 

 

  private loadCustomers(): void {

 

 

 

 

    this.props.spHttpClient

 

 

 

 

      .get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Customers')/items?$select=CustomerID,CustomerName`, SPHttpClient.configurations.v1)

 

 

 

 

      .then((response: SPHttpClientResponse) => {

 

 

 

 

        if (response.ok) {

 

 

 

 

          response.json().then((data) => {

 

 

 

 

            this.setState({ customers: data.value });

 

 

 

 

          });

 

 

 

 

        }

 

 

 

 

      });

 

 

 

 

  }

 

 

 

 

 

 

 

  // Inside the loadProducts() method:

 

 

 

 

  private loadProducts(): void {

 

 

 

 

    this.props.spHttpClient

 

 

 

 

      .get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Products')/items?$select=ProductID,ProductName,ProductType,ProductExpireDate,ProductUnitPrice`, SPHttpClient.configurations.v1)

 

 

 

 

      .then((response: SPHttpClientResponse) => {

 

 

 

 

        if (response.ok) {

 

 

 

 

          response.json().then((data) => {

 

 

 

 

            this.setState({ products: data.value });

 

 

 

 

          });

 

 

 

 

        }

 

 

 

 

      });

 

 

 

 

  }

 

 

 

 

  // Inside the loadOrders() method:

 

 

 

 

 

 

 

  private loadOrders(): void {

 

 

 

 

    this.props.spHttpClient

 

 

 

 

      .get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Orders')/items?$select=ID,OrderID,CustomerID,ProductID,UnitsSold,UnitPrice,SaleValue,OrderStatus`, SPHttpClient.configurations.v1)

 

 

 

 

      .then((response: SPHttpClientResponse) => {

 

 

 

 

 

 

 

        if (response.ok) {

 

 

 

 

          response.json().then((data) => {

 

 

 

 

            const orders = data.value;

 

 

 

 

            this.setState({ orders, selectedOrderID: orders.length > 0 ? orders[0].OrderID : '' });

 

 

 

 

          });

 

 

 

 

        }

 

 

 

 

      });

 

 

 

 

  }

 

 

 

 

 

 

 

  private handleCustomerChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {

 

 

 

 

    const selectedCustomer = event.target.value;

 

 

 

 

    this.setState({ selectedCustomer });

 

 

 

 

 

 

 

  };

 

 

 

 

 

 

 

  private handleProductChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {

 

 

 

 

    const selectedProduct = event.target.value;

 

 

 

 

    const product = this.state.products.find((p) => p.ProductName === selectedProduct);

 

 

 

 

    if (product) {

 

 

 

 

      this.setState({

 

 

 

 

        selectedProduct,

 

 

 

 

        selectedProductType: product.ProductType,

 

 

 

 

        selectedExpiryDate: product.ProductExpireDate,

 

 

 

 

        selectedUnitPrice: product.ProductUnitPrice,

 

 

 

 

      });

 

 

 

 

    }

 

 

 

 

  };

 

 

 

 

 

 

 

  private handleunitsSoldChange = (event: React.ChangeEvent<HTMLInputElement>): void => {

 

 

 

 

    const unitsSold = Number(event.target.value);

 

 

 

 

    this.setState({ unitsSold });

 

 

 

 

 

 

 

  };

 

 

 

 

 

 

 

  private generateRandomOrderID(): string {

 

 

 

 

    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';

 

 

 

 

    const length = 6;

 

 

 

 

    let OrderID = '';

 

 

 

 

 

 

 

    for (let i = 0; i < length; i++) {

 

 

 

 

      const randomIndex = Math.floor(Math.random() * characters.length);

 

 

 

 

      OrderID += characters.charAt(randomIndex);

 

 

 

 

    }

 

 

 

 

    return OrderID;

 

 

 

 

 

 

 

  }

 

 

 

 

  //Add the entry in orders list

 

 

 

 

 

 

 

  private handleAddOrder = (): void => {

 

 

 

 

    const { selectedCustomer, selectedProduct, selectedUnitPrice, unitsSold } = this.state;

 

 

 

 

    const saleValue = selectedUnitPrice * unitsSold;

 

 

 

 

    const customerObject = this.state.customers.find((c) => c.CustomerName === selectedCustomer)

 

 

 

 

    const productObject = this.state.products.find((p) => p.ProductName === selectedProduct);

 

 

 

 

    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Orders')/items`, SPHttpClient.configurations.v1, {

 

 

 

 

 

 

 

      body: JSON.stringify({

 

 

 

 

        OrderID: this.generateRandomOrderID(),

 

 

 

 

        CustomerID: customerObject.CustomerID,

 

 

 

 

        ProductID: productObject.ProductID,

 

 

 

 

        UnitsSold: unitsSold,

 

 

 

 

        UnitPrice: selectedUnitPrice,

 

 

 

 

        SaleValue: saleValue,

 

 

 

 

        OrderStatus: 'pending'

 

 

 

 

 

 

 

      }),

 

 

 

 

 

 

 

    }).then((response: SPHttpClientResponse) => {

 

 

 

 

 

 

 

      if (response.ok) {

 

 

 

 

        alert('Order added successfully');

 

 

 

 

      } else {

 

 

 

 

        response.json().then((data) => {

 

 

 

 

          alert('Error adding order: ' + JSON.stringify(data));

 

 

 

 

 

 

 

        });

 

 

 

 

      }

 

 

 

 

    })

 

 

 

 

 

 

 

      .catch((error) => {

 

 

 

 

        alert('Error adding order: ' + error);

 

 

 

 

 

 

 

      });

 

 

 

 

    // Clear the form values

 

 

 

 

 

 

 

    this.setState({

 

 

 

 

      selectedCustomer: '',

 

 

 

 

      selectedProduct: '',

 

 

 

 

      selectedProductType: '',

 

 

 

 

      selectedExpiryDate: '',

 

 

 

 

      selectedUnitPrice: 0,

 

 

 

 

      unitsSold: 0,

 

 

 

 

 

 

 

    });

 

 

 

 

    // Reload the orders

 

 

 

 

    this.loadOrders();

 

 

 

 

 

 

 

  };

 

 

 

 

  //Edit the order List

 

 

 

 

 

 

 

  private handleEditOrder = (): void => {

 

 

 

 

    const { selectedOrderID, selectedCustomer, selectedProduct, selectedUnitPrice, unitsSold } = this.state;

 

 

 

 

    const saleValue = selectedUnitPrice * unitsSold;

 

 

 

 

    const customerObject = this.state.customers.find((c) => c.CustomerName === selectedCustomer);

 

 

 

 

    const productObject = this.state.products.find((p) => p.ProductName === selectedProduct);

 

 

 

 

 

 

 

 

    const headers: any = {

 

 

 

 

      "Content-Type": "application/json",

 

 

 

 

      "Accept": "application/json",

 

 

 

 

      "X-HTTP-Method": "MERGE",

 

 

 

 

      "IF-MATCH": "*"

 

 

 

 

 

 

 

    };

 

 

 

 

 

 

 

 

    const spHttpClientOptions: ISPHttpClientOptions = {

 

 

 

 

 

 

 

      "headers": headers,

 

 

 

 

      "body": JSON.stringify({

 

 

 

 

        OrderID: selectedOrderID,

 

 

 

 

        CustomerID: customerObject.CustomerID,

 

 

 

 

        ProductID: productObject.ProductID,

 

 

 

 

        UnitsSold: unitsSold,

 

 

 

 

        UnitPrice: selectedUnitPrice,

 

 

 

 

        SaleValue: saleValue,

 

 

 

 

        OrderStatus: 'pending'

 

 

 

 

 

 

 

      })

 

 

 

 

 

 

 

    };

 

 

 

 

 

 

 

    const orderObject = this.state.orders.find((o) => o.OrderID === selectedOrderID)

 

 

 

 

    const url: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('Orders')/items('${orderObject.ID}')`;

 

 

 

 

 

 

 

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)

 

 

 

 

      .then((response: SPHttpClientResponse) => {

 

 

 

 

 

 

 

        if (response.ok) {

 

 

 

 

          alert('Order edited successfully');

 

 

 

 

        } else {

 

 

 

 

 

 

 

          response.json().then((data) => {

 

 

 

 

            alert('Error editing order: ' + JSON.stringify(data));

 

 

 

 

 

 

 

          });

 

 

 

 

        }

 

 

 

 

      })

 

 

 

 

 

 

 

      .catch((error) => {

 

 

 

 

        alert('Error editing order: ' + error);

 

 

 

 

 

 

 

      });

 

 

 

 

 

 

 

    // Clear the form values

 

 

 

 

 

 

 

    this.setState({

 

 

 

 

 

 

 

      selectedCustomer: '',

 

 

 

 

      selectedProduct: '',

 

 

 

 

      selectedProductType: '',

 

 

 

 

      selectedExpiryDate: '',

 

 

 

 

      selectedUnitPrice: 0,

 

 

 

 

      unitsSold: 0,

 

 

 

 

 

 

 

    });

 

 

 

 

 

 

 

    // Reload the orders

 

 

 

 

    this.loadOrders();

 

 

 

 

 

 

 

  };

 

 

 

 

 

 

 

  // Perform the necessary logic to delete the order with the provided order ID

 

 

 

 

  private handleDeleteOrder = (): void => {

 

 

 

 

    const { selectedOrderID } = this.state;

 

 

 

 

    const orderObject = this.state.orders.find((o) => o.OrderID === selectedOrderID)

 

 

 

 

    const url: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('Orders')/items('${orderObject.ID}')`;

 

 

 

 

    const headers: any = {

 

 

 

 

 

 

 

      "X-HTTP-Method": "DELETE",

 

 

 

 

      "IF-MATCH": "*",

 

 

 

 

      "Content-Type": "application/json"

 

 

 

 

 

 

 

    };

 

 

 

 

 

 

 

    const spHttpClientOptions: ISPHttpClientOptions = {

 

 

 

 

 

 

 

      "headers": headers

 

 

 

 

 

 

 

    };

 

 

 

 

 

 

 

    this.props.spHttpClient

 

 

 

 

 

 

 

      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)

 

 

 

 

      .then((response: SPHttpClientResponse) => {

 

 

 

 

 

 

 

        if (response.ok) {

 

 

 

 

          alert('Order deleted successfully');

 

 

 

 

 

 

 

          // Reload the orders

 

 

 

 

 

 

 

          this.loadOrders();

 

 

 

 

 

 

 

        } else {

 

 

 

 

          response.json().then((data) => {

 

 

 

 

            alert('Error deleting order: ' + JSON.stringify(data));

 

 

 

 

 

 

 

          });

 

 

 

 

 

 

 

        }

 

 

 

 

 

 

 

      })

 

 

 

 

 

 

 

      .catch((error) => {

 

 

 

 

        alert('Error deleting order ctach error: ' + error);

 

 

 

 

 

 

 

      });

 

 

 

 

 

 

 

 

    // Clear the form values

 

 

 

 

 

 

 

    this.setState({

 

 

 

 

 

 

 

      selectedCustomer: '',

 

 

 

 

      selectedProduct: '',

 

 

 

 

      selectedProductType: '',

 

 

 

 

      selectedExpiryDate: '',

 

 

 

 

      selectedUnitPrice: 0,

 

 

 

 

      unitsSold: 0,

 

 

 

 

 

 

 

    });

 

 

 

 

 

 

 

    // Reload the orders

 

 

 

 

 

 

 

    this.loadOrders();

 

 

 

 

 

 

 

  };

 

 

 

  private handleReset = (): void => {

 

 

    // Clear the form values

 

 

    this.setState({

 

 

      selectedCustomer: '',

 

 

 

 

      selectedProduct: '',

 

 

 

 

      selectedProductType: '',

 

 

 

 

      selectedExpiryDate: '',

 

 

 

 

      selectedUnitPrice: 0,

 

 

 

 

      unitsSold: 0,

 

 

 

 

 

 

 

    });

 

 

 

 

 

 

 

  };

 

 

 

  public render(): React.ReactElement<ISalesProps> {

 

 

    return (

 

 

 

      <div className={styles.form}>

 

 

 

 

        <h2>Order Entries</h2>

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <label>Customer Name:</label>

 

 

 

 

 

 

 

          <select value={this.state.selectedCustomer} onChange={this.handleCustomerChange}>

 

 

 

 

            <option value="">Select Customer</option>

 

 

 

 

            {this.state.customers.map((customer) => (

 

 

 

 

 

 

 

              <option key={customer.CustomerID} value={customer.CustomerName}>

 

 

 

 

                {customer.CustomerName}

 

 

 

 

              </option>

 

 

 

 

 

 

 

            ))}

 

 

 

 

 

 

 

          </select>

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <label>Product Name:</label>

 

 

 

 

 

 

 

          <select value={this.state.selectedProduct} onChange={this.handleProductChange}>

 

 

 

 

            <option value="">Select Product</option>

 

 

 

 

            {this.state.products.map((product) => (

 

 

 

 

 

 

 

              <option key={product.ProductID} value={product.ProductName}>

 

 

 

 

                {product.ProductName}

 

 

 

 

              </option>

 

 

 

 

 

 

 

            ))}

 

 

 

 

 

 

 

          </select>

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <label>Product Type:</label>

 

 

 

 

          <input type="text" value={this.state.selectedProductType} readOnly /> {/* Use the readOnly attribute */}

 

 

 

 

 

 

 

        </div>

 

 

 

        <br></br>

 

        <div>

 

 

          <label>Product Expiry Date:</label>

 

 

          <input type="text" value={this.state.selectedExpiryDate} readOnly /> {/* Use the readOnly attribute */}

 

        </div>

 

        <br></br>

 

        <div>

 

 

          <label>Product Unit Value:</label>

 

 

          <input type="number" value={this.state.selectedUnitPrice} readOnly /> {/* Use the readOnly attribute */}

 

 

 

        </div>

 

        <br></br>

 

        <div>

 

          <label>Number of Units:</label>

 

 

          <input

 

 

 

            type="number"

 

            value={this.state.unitsSold}

 

 

            onChange={this.handleunitsSoldChange}

 

 

            min="0"

 

 

 

 

            step="1"

 

 

 

 

 

 

 

          />

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <label>Sale Value:</label>

 

 

 

 

          <input type="text" value={this.state.selectedUnitPrice * this.state.unitsSold} readOnly /> {/* Use the readOnly attribute */}

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <label>Order ID:</label>

 

 

 

 

          <input

 

 

 

 

 

 

 

            type="text"

 

 

 

 

            value={this.state.selectedOrderID}

 

 

 

 

            onChange={(event) => this.setState({ selectedOrderID: event.target.value })}

 

 

 

 

 

 

 

          />

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        <div>

 

 

 

 

 

 

 

          <button className={styles.button} onClick={this.handleAddOrder}>Add Order</button> {/* Button for adding entry the Orders List */}

 

 

 

 

          <button className={styles.button} onClick={this.handleEditOrder}>Edit Order</button> {/* Button for editting the Orders List */}

 

 

 

 

          <button className={styles.button} onClick={this.handleDeleteOrder}>Delete Order</button> {/* Button for Deleting a entry in the Orders List */}

 

 

 

 

          <button className={styles.button} onClick={this.handleReset}>Reset</button> {/* Button for resetting the Form */}

 

 

 

 

 

 

 

        </div>

 

 

 

 

 

 

 

        <br></br>

 

 

 

 

 

 

 

        {/* Render the list of orders */}

 

 

 

 

 

 

 

        <h2>Orders</h2>

 

 

 

 

 

 

 

        <table>

 

 

 

 

 

 

 

          <thead>

 

 

 

 

 

 

 

            <tr>

 

 

 

 

 

 

 

              <th>Customer Id</th>

 

 

 

 

              <th >Product Id</th>

 

 

 

 

              <th >Product Unit Value</th>

 

 

 

 

              <th >Number of Units</th>

 

 

 

 

              <th >Sale Value</th>

 

 

 

 

 

 

 

            </tr>

 

 

 

 

 

 

 

          </thead>

 

 

 

 

 

 

 

          <tbody>

 

 

 

 

 

 

 

            {this.state.orders.map((order) => (

 

 

 

 

 

 

 

              <tr key={order.OrderID}>

 

 

 

 

 

 

 

                <td >{order.CustomerID}</td>

 

 

 

 

                <td >{order.ProductID}</td>

 

 

 

 

                <td >{order.UnitPrice}</td>

 

 

 

 

                <td >{order.UnitsSold}</td>

 

 

 

 

                <td >{order.SaleValue}</td>

 

 

 

 

 

 

 

              </tr>

 

 

 

 

 

 

 

            ))}

 

 

 

 

          </tbody>

 

 

 

 

        </table>

 

 

 

 

      </div>

 

 

 

 

 

 

 

    );

 

 

 

 

  }

 

 

 

 

 

 

 

}