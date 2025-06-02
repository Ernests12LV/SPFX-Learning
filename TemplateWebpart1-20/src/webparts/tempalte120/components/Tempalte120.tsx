import * as React from "react";
import styles from "./Tempalte120.module.scss";
import type { ITempalte120Props } from "./ITempalte120Props";
import { escape } from "@microsoft/sp-lodash-subset";

export default class Tempalte120 extends React.Component<ITempalte120Props> {
  public render(): React.ReactElement<ITempalte120Props> {
    const {
      productName,
      productDescription,
      productCost,
      quantity,
      billAmount,
      discount,
      netBillAmount,
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      lists,
      isCertified,
      rating,
      category,
      deliveryOption,
      features,
      paymentMethod,
      colorScheme,
      listName,
      listDescription,
    } = this.props;

    return (
      <section
        className={`${styles.tempalte120} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div className={styles.productDetails}>
          <h3>Product Details</h3>
          <div>
            <label>Product Name:</label>
            <span>{productName}</span>
          </div>
          <div>
            <label>Description:</label>
            <span>{productDescription}</span>
          </div>
          <div>
            <label>Cost:</label>
            <span>${productCost}</span>
          </div>
          <div>
            <label>Quantity:</label>
            <span>{quantity}</span>
          </div>
          <div>
            <label>Bill Amount:</label>
            <span>${billAmount}</span>
          </div>
          <div>
            <label>Discount:</label>
            <span>${discount}</span>
          </div>
          <div>
            <label>Net Bill Amount:</label>
            <span>${netBillAmount}</span>
          </div>
          <div>
            <label>Certification Status:</label>
            <span>{isCertified ? "Certified" : "Not Certified"}</span>
          </div>
          <div>
            <label>Product Rating:</label>
            <span>{rating}/10</span>
          </div>
          <div>
            <label>Category:</label>
            <span>{category.charAt(0).toUpperCase() + category.slice(1)}</span>
          </div>
          <div>
            <label>Delivery Option:</label>
            <span>{deliveryOption}</span>
          </div>
          <div>
            <label>Selected Features:</label>
            <span>{features.join(", ")}</span>
          </div>
          <div>
            <label>Payment Method:</label>
            <span>{paymentMethod}</span>
          </div>
          <div>
            <label>Color Scheme:</label>
            <span>{colorScheme}</span>
          </div>
        </div>
        <div>
          <h4>SharePoint Lists:</h4>
          <ul>
            {lists.map((list) => (
              <li key={list.Id}>
                {list.Title} ({list.Id})
              </li>
            ))}
          </ul>
        </div>
        <div className={styles.listManagement}>
          <h3>List Management</h3>
          <div>
            <label>List Name:</label>
            <span>{listName}</span>
          </div>
          <div>
            <label>Description:</label>
            <span>{listDescription}</span>
          </div>
        </div>
      </section>
    );
  }
}
